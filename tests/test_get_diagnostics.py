from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from trace_parser import gripper_reply, read_trace
from trace_store import TraceStore, Cancelled


HEADER = 'creation de la trace de ... à 00:00:01 Le 13/10/2026\n'


def ticket(ident='1052', ref='REF', sn='SN', lot='LOT'):
    return f' 13|13:37:15.8 RefProduit={ref}<Cr>SN={sn}<Cr>CodeLot={lot}<Cr>ModBusId={ident}<Cr>\n'


def get(ref='REF', extra=''):
    return f'G 13|13:37:18.9 GRIPPER GET 1x boîte(s): Code:{ref} l:0 h:0 L:0mm face:1 $0101 {extra}\n'


def pute(ref='REF'):
    return f'G 13|13:46:49.6 GRIPPER PUTE 1x boîte(s): Code:{ref} l:92 h:17 L:132mm face:2 D23\n'


FAILURE = "r [Tgripper = 5s rep ='Err NOK (-32092)']\nG 13|13:37:23.9 Err: reponse gripper GET: Err NOK (-32092)\n"
OK = "r [Tgripper = 5s rep ='OK']\n"


class GetDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.path = self.root/'traces'/'R2'/'TracOmega.txt'
        self.path.parent.mkdir(parents=True)
        self.store = TraceStore(self.root/'cache.sqlite3', directories=[str(self.root/'traces')])

    def publish(self, body):
        self.path.write_text(HEADER+body, encoding='utf8')
        self.store.sync([str(self.path)])
        self.ref, hits = next((p, hits) for p, hits in self.store.search('ModBusId=1052', True) if not p.historical)
        self.number = hits[0][0]
        return self.store.get_for_occurrence(self.ref, self.number)

    def test_failed_get_does_not_claim_later_deposit_and_keeps_proofs(self):
        result = self.publish(ticket()+get()+FAILURE+ticket('1060', sn='OTHER')+get()+OK+pute())
        self.assertEqual(result['status'], 'failed')
        self.assertIn('-32092', result['message'])
        self.assertEqual(len(result['responses']), 2)
        self.assertIsNotNone(self.store.details(self.ref, self.number)['pute'])
        self.assertIsNone(result['pute'])
        self.assertTrue(all(r['path']==self.ref and r['number']>self.number for r in result['responses']))

    def test_targeted_diagnostics_equal_complete_stream_with_interleaved_references(self):
        from get_diagnostics import _analyze, diagnostics
        body = ticket()+get()+FAILURE
        for i in range(60):
            ref = 'REF' if i%4==0 else f'OTHER{i%5}'
            context = f"b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:{10000+i}({ref})) en $0101\n"
            body += ticket(str(2000+i), ref=ref, sn=str(i))+context+get(ref)
            body += FAILURE if i%3==0 else OK
            body += 'G GRIPPER REF\n'+OK if i%2 else 'G GRIPPER RESET\n'+OK
            body += get('UNREQUESTED')+OK
            if i%3==0:
                body += context+get(ref)+OK
            body += pute(ref)
        self.publish(body)
        rows = self.store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE f.path=? ORDER BY l.number''', (self.ref,))
        expected = _analyze(self.store, rows, None)
        wanted = [row for row in rows if row.get('modbus') is not None and row.get('ref') in ('REF','OTHER2')]
        self.store._analysis_cache.clear()
        actual = diagnostics(self.store, wanted, self.store.analysis_revision())
        self.assertEqual(actual, {(row['path'], row['number']): expected[row['number']] for row in wanted})

    def test_concurrent_local_and_get_analysis_share_work_and_wait_is_cancellable(self):
        from concurrent.futures import ThreadPoolExecutor
        import get_diagnostics
        context = "b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:142509(REF)) en $0101\n"
        self.publish(ticket()+context+get()+FAILURE+context+get()+OK
                     +'B 13|13:37:20.0 IDPK pour REF: l=100 H=50 L=195\n')
        self.store._analysis_cache.clear()
        entered, release, stop = threading.Event(), threading.Event(), threading.Event()
        original = get_diagnostics._targeted_rows
        def slow(*args, **kwargs):
            entered.set()
            if not release.wait(3):
                raise AssertionError('Analysis gate timed out')
            return original(*args, **kwargs)
        with patch.object(get_diagnostics, '_targeted_rows', side_effect=slow) as reads, ThreadPoolExecutor(3) as pool:
            first = pool.submit(self.store.current_measurement_for_occurrence, self.ref, self.number)
            try:
                self.assertTrue(entered.wait(2))
                waiting = pool.submit(self.store.current_measurement_for_occurrence, self.ref, self.number, stop)
                stop.set()
                with self.assertRaises(Cancelled):
                    waiting.result(timeout=2)
                second = pool.submit(self.store.get_for_occurrence, self.ref, self.number)
            finally:
                release.set()
            self.assertIsNotNone(first.result(timeout=3)['current'])
            self.assertEqual(second.result(timeout=3)['status'], 'recovered')
            self.assertEqual(reads.call_count, 1)
        self.assertFalse(self.store._analysis_cache.running)

    def test_dense_reference_history_streams_once_and_matches_complete_analysis(self):
        import get_diagnostics
        body = ticket()+get()+OK
        for i in range(80):
            ref = 'OTHER' if i%5==0 else 'REF'
            body += ticket(str(3000+i), ref=ref, sn=str(i))+get(ref)
            body += FAILURE if i%3==0 else OK
            body += 'G GRIPPER EMPTY\n'+OK+pute(ref)
        self.publish(body)
        rows = self.store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE f.path=? ORDER BY l.number''', (self.ref,))
        expected = get_diagnostics._analyze(self.store, rows, None)
        wanted = [row for row in rows if row.get('modbus') is not None and row.get('ref')=='REF']
        self.store._analysis_cache.clear()
        with patch.object(get_diagnostics, '_dense_rows', wraps=get_diagnostics._dense_rows) as dense, \
             patch.object(get_diagnostics, '_neighbor_command', side_effect=AssertionError('Per-GET query in dense stream')):
            actual = get_diagnostics.diagnostics(self.store, wanted, self.store.analysis_revision())
        self.assertEqual(dense.call_count, 1)
        self.assertEqual(actual, {(row['path'], row['number']): expected[row['number']] for row in wanted})

    def test_filtered_stream_keeps_long_retry_and_case_insensitive_replies(self):
        context = "b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:142509(REF)) en $0101\n"
        noise = 'B 13|13:37:19.0 routine sans evenement utile\n' * 2500
        self.publish(ticket()+context+noise+get()+FAILURE+'G GRIPPER EMPTY\n'+OK
                     +context+noise+get()+OK.replace('Tgripper', 'TGRIPPER')+pute())
        self.store._analysis_cache.clear()
        original = self.store._rows
        transferred = []
        def counted(*args, **kwargs):
            result = original(*args, **kwargs)
            transferred.extend(result)
            return result
        with patch.object(self.store, '_rows', side_effect=counted):
            result = self.store.get_for_occurrence(self.ref, self.number)
        self.assertEqual(result['status'], 'recovered')
        self.assertLess(len(transferred), 20)
        self.assertTrue(all('routine sans' not in row['text'] for row in transferred))

    def test_unknown_is_not_success_and_other_commands_stop_reply_window(self):
        result = self.publish(ticket()+get()+'G 13|13:38:00.0 GRIPPER REF\n'+OK+pute())
        self.assertEqual(result['status'], 'unknown')
        self.assertFalse(result['failed'])
        self.assertIsNone(result['pute'])
        result = self.publish(ticket()+get()+'G 13|13:38:00.0 GRIPPER EMPTY\n'+FAILURE)
        self.assertEqual(result['status'], 'unknown')
        self.assertFalse(result['responses'])

    def test_confirmed_retry_keeps_failure_but_plain_get_cannot_inherit(self):
        explicit = 'RefProduit=REF<Cr>SN=SN<Cr>CodeLot=LOT<Cr>ModBusId=1052<Cr>'
        result = self.publish(ticket()+get()+FAILURE+get(extra=explicit)+OK+pute())
        self.assertEqual(result['status'], 'recovered')
        self.assertTrue(result['failed'])
        self.assertEqual(len(result['attempts']), 2)
        self.assertIsNotNone(result['pute'])
        result = self.publish(ticket()+get()+FAILURE+get()+OK+pute())
        self.assertEqual(result['status'], 'failed')
        self.assertIsNone(result['pute'])

    def test_success_missing_get_and_reused_id_remain_separate(self):
        result = self.publish(ticket()+get()+OK+pute()+ticket(sn='SECOND')+get()+FAILURE)
        self.assertEqual(result['status'], 'success')
        self.assertIsNotNone(result['pute'])
        numbers = self.store.search('ModBusId=1052', True)[0][1]
        self.assertEqual(self.store.get_for_occurrence(self.ref, numbers[1][0])['status'], 'failed')
        result = self.publish(ticket()+ticket('1060')+get()+OK+pute())
        self.assertEqual(result['status'], 'missing')

    def test_revision_cache_deleted_generations_and_cancellation(self):
        result = self.publish(ticket()+get()+FAILURE)
        old_ref, old_number = self.ref, self.number
        with patch.object(self.store, '_rows', side_effect=AssertionError('Unchanged diagnostic read')):
            self.assertEqual(self.store.get_for_occurrence(self.ref, self.number), result)
        self.publish(ticket()+get()+OK+pute())
        rows = self.store.search('ModBusId=1052', True)
        self.assertEqual(len(rows), 2)
        self.assertEqual(self.store.get_for_occurrence(old_ref, old_number)['status'], 'failed')
        self.path.unlink(); self.store.sync([])
        self.assertEqual(self.store.get_for_occurrence(old_ref, old_number)['status'], 'failed')
        stop = threading.Event(); stop.set()
        with self.assertRaises(Cancelled):
            self.store.get_for_occurrence(old_ref, old_number, stop)

    def test_reply_variants(self):
        self.assertEqual(gripper_reply("r [Tgripper=5s rep= 'OK']")['status'], 'success')
        self.assertEqual(gripper_reply('G Err: réponse gripper GET: Err NOK (-32092)')['code'], '-32092')
        self.assertIsNone(gripper_reply('ERREUR deltaW=196'))

    def test_internal_box_retry_survives_reset_and_rejects_another_box(self):
        def context(box):
            return f"b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:{box}(REF)) en $0101\n"
        body = ticket()+context('142509')+get()+FAILURE+'G GRIPPER REF\n'+OK
        result = self.publish(body+context('142509')+get()+OK+pute())
        self.assertEqual(result['status'], 'recovered')
        self.assertEqual(len(result['attempts']), 2)
        self.assertEqual(result['attempts'][1]['status'], 'success')
        result = self.publish(body+context('999999')+get()+OK+pute())
        self.assertEqual(result['status'], 'failed')
        result = self.publish(body+ticket('1060', sn='OTHER')+context('142509')+get()+OK+pute())
        self.assertEqual(result['status'], 'failed')

    def test_idpk_after_confirmed_retry_belongs_to_original_ticket(self):
        context = "b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:142509(REF)) en $0101\n"
        measure = 'B 13|13:46:48.0 IDPK pour REF : l=104 H=49 L=195\n'
        body = ticket()+context+get()+FAILURE+'G GRIPPER REF\n'+OK
        body += context+get()+FAILURE+context+get()+OK+measure+pute()
        self.publish(body)
        result = self.store.comparison_for_occurrence(self.ref, self.number)
        self.assertIsNotNone(result['current'])
        self.assertEqual(result['current']['text'], measure.rstrip())
        self.assertEqual(result['current']['path'], self.ref)
        self.assertEqual(result['status'], 'no_previous')
        local = self.store.current_measurement_for_occurrence(self.ref, self.number)
        self.assertEqual(local['current'], result['current'])
        self.assertEqual(local['revision'], result['revision'])
        # Without the explicit box identity, the successful GET cannot inherit.
        self.publish(ticket()+context+get()+FAILURE+get()+OK+measure+pute())
        self.assertIsNone(self.store.comparison_for_occurrence(self.ref, self.number)['current'])
        # A new ticket invalidates the former box ownership, even with same IDb.
        self.publish(ticket()+context+get()+FAILURE+ticket('1060', sn='OTHER')+context+get()+OK+measure+pute())
        self.assertIsNone(self.store.comparison_for_occurrence(self.ref, self.number)['current'])

    def test_reported_701_retry_read_only(self):
        path = Path('traces/trace r2/TracOMEGA.txt')
        if not path.exists():
            self.skipTest('Local trace absent')
        lines = read_trace(path).splitlines()
        start = next(i for i, line in enumerate(lines) if 'ModBusId=701<Cr>' in line and 'A0117192AC' in line)
        end = next(i for i in range(start+1, len(lines)) if 'GRIPPER PUTE' in lines[i] and 'Code:A0117192AC ' in lines[i])
        self.path.write_text(HEADER+'\n'.join(lines[start:end+1])+'\n', encoding='utf8')
        self.store.sync([str(self.path)])
        ref, hits = self.store.search('ModBusId=701', True)[0]
        result = self.store.get_for_occurrence(ref, hits[0][0])
        self.assertEqual(result['status'], 'recovered')
        self.assertEqual(len(result['attempts']), 2)
        self.assertIn('11:22:39.6', result['attempts'][1]['get']['text'])
        self.assertEqual(result['attempts'][1]['status'], 'success')

    def test_local_reported_case_read_only(self):
        path = Path('traces/trace r2/TracOmega029.old')
        if not path.exists():
            self.skipTest('Local trace absent')
        lines = read_trace(path).splitlines()
        start = next(i for i, line in enumerate(lines) if 'ModBusId=1052<Cr>' in line and '261012.095' in line)
        end = next(i for i in range(start+1, len(lines)) if 'GRIPPER PUTE' in lines[i] and 'Code:A0110543AC ' in lines[i])
        result = self.publish('\n'.join(lines[start:end+1])+'\n')
        self.assertEqual(result['status'], 'recovered')
        self.assertIn('-32092', result['attempts'][0]['message'])
        self.assertIsNotNone(result['pute'])
        detail = self.store.details(self.ref, self.number)
        self.assertEqual(detail['pute']['text'], lines[end])

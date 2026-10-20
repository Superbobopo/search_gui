from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from trace_store import TraceStore, Cancelled
from tests.test_get_diagnostics import HEADER,ticket,get,FAILURE,OK,pute


class GetErrorsTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory(); self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.directory = self.root/'traces'
        self.store = TraceStore(self.root/'cache.sqlite3',directories=[str(self.directory)])
        self.files = []

    def publish(self,body,robot='R1'):
        path = self.directory/robot/'TracOmega.txt'
        path.parent.mkdir(parents=True,exist_ok=True)
        path.write_text(HEADER+body,encoding='utf8')
        if str(path) not in self.files: self.files.append(str(path))
        self.store.sync(self.files)
        return path

    def test_failure_responses_group_by_command_and_other_operations_do_not_leak(self):
        self.publish(ticket()+get()+FAILURE+OK+pute()+FAILURE.splitlines()[0]+'\n'
                     +'G GRIPPER REF\n'+OK+'G GRIPPER EMPTY\n'+OK+get('NO_REPLY'))
        with patch.object(self.store,'comparison_for_occurrence',side_effect=AssertionError('No IDPK')):
            result = self.store.get_errors()
        self.assertEqual(result['revision'],self.store.analysis_revision())
        self.assertEqual(len(result['errors']),1)
        row = result['errors'][0]
        self.assertEqual((row['ref'],row['modbus'],row['sn'],row['lot']),('REF','1052','SN','LOT'))
        self.assertEqual((row['robot'],row['face'],row['location']),('R1','1','$0101'))
        self.assertEqual(len(row['responses']),2)
        self.assertEqual(row['code'],'-32092')
        self.assertTrue(row['error']['reply']['explicit_get'])

    def test_success_or_missing_reply_is_not_a_failure(self):
        self.publish(ticket()+get()+OK+get('UNKNOWN'))
        self.assertEqual(self.store.get_errors()['errors'],[])

    def test_same_positions_in_different_robots_and_full_codes_remain_distinct(self):
        self.publish(ticket()+get()+FAILURE+get('SHELF').replace('$0101','F32')+FAILURE
                     +get('LONG').replace('$0101','$01010')+FAILURE)
        self.publish(ticket()+get()+FAILURE,robot='R2')
        rows = self.store.get_errors()['errors']
        self.assertEqual({(r['robot'],r['location']) for r in rows},
                         {('R1','$0101'),('R2','$0101'),('R1','F32'),('R1','$01010')})
        self.assertTrue(all(r['modbus'] is None for r in rows if r['location']!='$0101'))
        self.assertEqual(len({r['path'] for r in rows}),2)

    def test_explicit_unattached_error_is_kept_but_generic_other_reply_is_not(self):
        self.publish(FAILURE+'G GRIPPER PUTE 1x boîte(s): Code:OTHER face:2 F32\n'+FAILURE)
        rows = self.store.get_errors()['errors']
        self.assertEqual(len(rows),2)
        self.assertTrue(all(r['get'] is None and r['location'] is None and r['ref'] is None for r in rows))
        self.assertTrue(all(r['error']['reply']['explicit_get'] for r in rows))

    def test_retry_requires_identity_and_reused_id_keeps_new_ticket(self):
        context = "b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:142509(REF)) en $0101\n"
        self.publish(ticket()+context+get()+FAILURE+context+get()+FAILURE
                     +get()+FAILURE+ticket(sn='OTHER',lot='OTHERLOT')+get()+FAILURE)
        rows = sorted(self.store.get_errors()['errors'],key=lambda r:r['number'])
        self.assertEqual(len(rows),4)
        self.assertEqual([r['sn'] for r in rows],['SN','SN',None,'OTHER'])
        self.assertEqual(rows[-1]['lot'],'OTHERLOT')

    def test_generations_and_revision_are_preserved_and_cancel_is_honored(self):
        self.publish(ticket()+get()+FAILURE)
        first = self.store.get_errors()
        self.publish(ticket(sn='NEW')+get()+FAILURE)
        second = self.store.get_errors()
        self.assertNotEqual(first['revision'],second['revision'])
        self.assertEqual(len({r['path'].version_id for r in second['errors']}),2)
        self.assertEqual({r['sn'] for r in second['errors']},{'SN','NEW'})
        stop = threading.Event(); stop.set()
        with self.assertRaises(Cancelled): self.store.get_errors(stop)

    def test_uppercase_reply_is_not_mistaken_for_a_new_get_command(self):
        self.publish(ticket()+get()+FAILURE.replace('gripper GET','GRIPPER GET'))
        row, = self.store.get_errors()['errors']
        self.assertEqual(len(row['responses']),2)
        diagnostic = self.store.get_for_occurrence(row['path'],1)
        self.assertTrue(diagnostic['failed'])

    def test_fts_limits_transferred_lines_and_does_not_diagnose_ordinary_history(self):
        self.publish(ticket()+get()+FAILURE+'\n'.join('b ordinary noise' for _ in range(5000)))
        counts = []; original = self.store._rows
        def rows(*args,**kw):
            value = original(*args,**kw); counts.append(len(value)); return value
        with patch.object(self.store,'_rows',side_effect=rows), \
             patch.object(self.store,'get_for_occurrence',side_effect=AssertionError('No retry needed')):
            row, = self.store.get_errors()['errors']
        self.assertLess(sum(counts),30)
        self.assertEqual(row['modbus'],'1052')

    def test_all_failure_words_and_accented_explicit_reply_are_fts_candidates(self):
        for value in ('Err NOK (-7)','error (-2)','erreur (-3)','NOK (-4)'):
            with self.subTest(value=value):
                self.publish(ticket()+get()+f"r [Tgripper rep='{value}']\n")
                rows = [r for r in self.store.get_errors()['errors'] if not r['path'].historical]
                self.assertEqual(len(rows),1)
                self.assertEqual(rows[0]['message'],value)
        self.publish(ticket()+get()+'G 13|13:37:23.9 Err: réponse gripper GET: Erreur (-7)\n')
        row, = (r for r in self.store.get_errors()['errors'] if not r['path'].historical)
        self.assertEqual(row['code'],'-7')

    def test_conflicting_command_identity_does_not_borrow_ticket_fields(self):
        self.publish(ticket()+get(extra='RefProduit=OTHER SN=OTHER_SN CodeLot=OTHER_LOT ModBusId=99')+FAILURE)
        row, = self.store.get_errors()['errors']
        self.assertEqual(row['ref'],'REF')
        self.assertTrue(all(row[key] is None for key in ('sn','lot','modbus')))

"""Confirmed failed GET attempts, seeded by FTS rather than all occurrences."""
import re

from get_diagnostics import _neighbor_command, add_reply
from trace_parser import GRIPPER_COMMAND_RE, gripper_reply, gripper_location, parse_fields
from trace_sources import robot_scope


def _identity(store, command, reference, revision, cancel):
    fields = parse_fields(command['text'])
    if reference and fields.get('RefProduit') not in (None,reference):
        return dict(ref=reference,sn=None,lot=None,modbus=None)
    reference = reference or fields.get('RefProduit')
    identity = dict(ref=reference, sn=fields.get('SN'), lot=fields.get('CodeLot'), modbus=fields.get('ModBusId'))
    # A loading ticket does not identify a subsequent shelf pick of the same ref.
    if not reference or not command.get('loading'):
        return identity
    tickets = store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
        WHERE f.path=? AND l.ref=? AND l.modbus IS NOT NULL AND l.number<?
        ORDER BY l.number DESC LIMIT 1''', (command['path'],reference,command['number']),cancel)
    if not tickets:
        return identity
    ticket = tickets[0]
    previous = store._rows('''SELECT l.number FROM lines l JOIN files f ON f.id=l.file_id
        WHERE f.path=? AND l.op_ref=? AND l.op='GET' AND l.number>? AND l.number<? LIMIT 1''',
        (command['path'],reference,ticket['number'],command['number']),cancel)
    if previous:
        # Only confirmed retries can reuse an already consumed ticket.
        diagnostic = store.get_for_occurrence(command['path'],ticket['number'],cancel)
        if not any(a['get']['number']==command['number'] for a in diagnostic['attempts']):
            return identity
    if any(fields.get(name) is not None and fields[name]!=ticket[key]
           for name,key in (('SN','sn'),('CodeLot','lot'),('ModBusId','modbus'))):
        return identity
    return dict(ref=reference,sn=ticket.get('sn'),lot=ticket.get('lot'),modbus=ticket.get('modbus'))


def get_errors(store, revision, cancel):
    errors, after = {}, 0
    while True:
        store._check_cancel(cancel)
        # Every failure recognized by gripper_reply contains ERR or NOK. MATCH is
        # made of three-character terms, compatible with the existing detail=none FTS.
        rows = store._rows('''SELECT l.*,f.path FROM lines_fts
            JOIN lines l ON l.id=lines_fts.rowid JOIN files f ON f.id=l.file_id
            WHERE lines_fts MATCH ? AND lines_fts.rowid>? AND f.kind='omega'
            AND (instr(l.normalized,'tgripper')>0 OR
                 (instr(l.normalized,'get')>0 AND
                  (instr(l.normalized,'reponse')>0 OR instr(l.normalized,'réponse')>0)))
            ORDER BY lines_fts.rowid LIMIT 500''',
            ('("err" OR "nok") AND "gri" AND "rip" AND "ipp" AND "ppe" AND "per"',after),cancel)
        if not rows:
            break
        for row in rows:
            store._check_cancel(cancel)
            reply = gripper_reply(row['text'])
            if not reply or reply['status']!='failed':
                continue
            command = _neighbor_command(store,row['path'],row['number'],-1,revision,cancel)
            operation = GRIPPER_COMMAND_RE.search(command['text']) if command else None
            if not operation or operation[1]!='GET':
                if not reply['explicit_get']:
                    continue
                command = None
            key = (row['path'], 'get' if command else 'orphan',command['number'] if command else row['number'])
            record = errors.get(key)
            if record is None:
                record = errors[key] = dict(path=row['path'],get=command,status='unknown',message='',responses=[])
            add_reply(record,row,reply)
        after = rows[-1]['id']
        if len(rows)<500:
            break
    for record in errors.values():
        store._check_cancel(cancel)
        command = record['get']
        reference = re.search(r'\bCode:\s*([^<>\s;]+)',command['text']) if command else None
        reference = reference[1] if reference else None
        record.update(_identity(store,command,reference,revision,cancel) if command else
                      dict(ref=None,sn=None,lot=None,modbus=None))
        record.update(gripper_location(command['text']) if command else dict(face=None,location=None))
        record['robot'] = robot_scope(record['path'])[1]
        # Prefer the explicit, timestamped GET error as the opening destination.
        record['error'] = next((r for r in record['responses'] if r['reply']['explicit_get']),record['responses'][0])
        record['time'] = record['error']['time']
        record['code'] = record['responses'][-1]['reply']['code']
        record['number'] = command['number'] if command else record['error']['number']
    return dict(revision=revision,errors=sorted(errors.values(),key=lambda r:(r['time'],str(r['path']),r['number']),reverse=True))

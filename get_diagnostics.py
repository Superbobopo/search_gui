"""Occurrence-scoped GET outcomes from indexed original lines only."""
from trace_parser import GRIPPER_COMMAND_RE, GRIPPER_RE, gripper_reply, parse_fields, get_box_context


def add_reply(command, row, reply):
    """Shared precedence: a later OK never erases an explicit failure."""
    command['responses'].append(dict(row, reply=reply))
    if command['status']!='failed' or reply['status']=='failed':
        command.update(status=reply['status'], message=reply['message'])


def _identity(row):
    return tuple(row.get(key) for key in ('ref', 'sn', 'lot', 'modbus'))


def _finish(record):
    attempts = record['attempts']
    failed = [a for a in attempts if a['status']=='failed']
    recovered = bool(failed and attempts[-1]['status']=='success')
    status = ('recovered' if recovered else 'failed') if failed else (
        attempts[-1]['status'] if attempts else 'missing')
    message = ('GET échoué puis réussi' if recovered else
               'GET échoué — '+failed[-1]['message'] if failed else
               'GET réussi' if status=='success' else
               'GET non trouvé pour cette occurrence' if status=='missing' else 'Résultat GET inconnu')
    return dict(status=status, failed=bool(failed), message=message, attempts=attempts,
                get=attempts[0]['get'] if attempts else None,
                pute=record.get('pute'), responses=[r for a in attempts for r in a['responses']])


def _neighbor_command(store, path, number, direction, revision, cancel):
    key = ('gripper_neighbor', revision, path, number, direction)
    cached = store._analysis_cache.get(key)
    if cached is not None:
        return cached[0]
    relation, ordering = ('>', 'ASC') if direction>0 else ('<', 'DESC')
    after = number
    while True:
        store._check_cancel(cancel)
        rows = store._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id '
            'WHERE f.path=? AND l.number '+relation+" ? AND instr(l.text,'GRIPPER')>0 "
            'ORDER BY l.number '+ordering+' LIMIT 2', (path, after), cancel)
        for row in rows:
            if GRIPPER_COMMAND_RE.search(row['text']) and not gripper_reply(row['text']):
                store._analysis_cache.put(key, (row,))
                return row
        if len(rows)<2:
            store._analysis_cache.put(key, (None,))
            return None
        after = rows[-1]['number']


def _dense_rows(store, path, references, cancel):
    """For dense histories, scan useful events once rather than query per GET."""
    marks = ','.join('?' for _ in references)
    after = -1
    while True:
        store._check_cancel(cancel)
        rows = store._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id '
            'WHERE f.path=? AND l.number>? AND ((l.modbus IS NOT NULL AND l.ref IN ('+marks+')) '
            "OR instr(l.text,'GRIPPER')>0 OR l.text LIKE '%tgripper%' "
            "OR l.text LIKE '%chargeboitesdansgrippermb%' "
            "OR (l.text LIKE '%gripper%' AND l.text LIKE '%GET%')) ORDER BY l.number LIMIT 1000",
            (path, after, *references), cancel)
        if not rows:
            return
        yield from rows
        after = rows[-1]['number']


def _targeted_rows(store, path, references, revision, cancel):
    marks = ','.join('?' for _ in references)
    base = store._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id '
        'WHERE f.path=? AND ((l.ref IN ('+marks+') AND l.modbus IS NOT NULL) '
        +"OR (l.op_ref IN ("+marks+") AND l.op IN ('GET','PUTE'))) ORDER BY l.number",
        (path, *references, *references), cancel)
    gets = sum(row.get('op')=='GET' for row in base)
    if gets>32:
        total = store._rows("SELECT count(*) AS n FROM lines WHERE file_id=(SELECT id FROM files WHERE path=?) AND op='GET'",
                            (path,), cancel)[0]['n']
        if gets*4>=total:
            return _dense_rows(store, path, references, cancel)
    entries = {row['number']: row for row in base}
    # REF/EMPTY invalidate the implicit retry identity, even between distant GETs.
    resets_key = ('gripper_resets', revision, path)
    resets = store._analysis_cache.get(resets_key)
    if resets is None:
        candidates = store._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id '
            "WHERE f.path=? AND l.op IS NULL AND instr(l.text,'GRIPPER')>0 "
            "AND (instr(l.text,'REF')>0 OR instr(l.text,'EMPTY')>0) ORDER BY l.number",
            (path,), cancel)
        resets = []
        for row in candidates:
            store._check_cancel(cancel)
            cmd = GRIPPER_COMMAND_RE.search(row['text'])
            if cmd and cmd[1] in ('REF', 'EMPTY'):
                resets.append(row)
        store._analysis_cache.put(resets_key, resets, len(resets))
    entries.update((row['number'], row) for row in resets)
    for command in base:
        store._check_cancel(cancel)
        if command.get('op')!='GET':
            continue
        number = command['number']
        previous = _neighbor_command(store, path, number, -1, revision, cancel)
        following = _neighbor_command(store, path, number, 1, revision, cancel)
        if previous is not None:
            entries[previous['number']] = previous
        lower = previous['number'] if previous else -1
        upper = following['number'] if following else 9223372036854775807
        rows = store._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id '
            "WHERE f.path=? AND ((l.number>? AND l.number<? AND l.text LIKE '%chargeboitesdansgrippermb%') "
            "OR (l.number>? AND l.number<? AND (l.text LIKE '%tgripper%' "
            "OR (l.text LIKE '%gripper%' AND l.text LIKE '%GET%')))) ORDER BY l.number",
            (path, lower, number, number, upper), cancel)
        entries.update((row['number'], row) for row in rows)
    return sorted(entries.values(), key=lambda row: row['number'])


def _analyze(store, rows, cancel, references=None):
    records, pending, current = {}, {}, {}
    box_owners = {}
    loading_context = None
    command = None
    for row in rows:
        store._check_cancel(cancel)
        box_context = get_box_context(row['text'])
        if box_context:
            loading_context = box_context
        reply = gripper_reply(row['text'])
        cmd = None if reply else GRIPPER_COMMAND_RE.search(row['text'])
        if cmd:
            command = None  # Every command terminates the previous reply window.
            context, loading_context = loading_context, None
            if cmd[1] in ('REF', 'EMPTY'):
                current.clear()
            op = GRIPPER_RE.search(row['text'])
            if not op:
                continue
            ref = op[2]
            if references is not None and ref not in references:
                continue
            if op[1]=='GET':
                ticket = pending.pop(ref, None)
                record = records.get(ticket['number']) if ticket else None
                previous = current.pop(ref, None)
                box_key = context if context and context[0]==ref else None
                owner = box_owners.get(box_key)
                if record is None and owner and owner['attempts'][-1]['status']=='failed':
                    record = owner
                # A retry needs explicit complete identity on its GET line.
                # A plain GET without a fresh ticket is never enough.
                if record is None and previous and previous['attempts'][-1]['status']=='failed':
                    fields = parse_fields(row['text'])
                    identity = tuple(fields.get(k) for k in ('RefProduit','SN','CodeLot','ModBusId'))
                    if all(identity) and identity==_identity(previous['ticket']):
                        record = previous
                if record is not None and row['loading']:
                    command = dict(get=row, status='unknown', message='', responses=[])
                    record['attempts'].append(command)
                    current[ref] = record
                    if box_key:
                        box_owners[box_key] = record
            elif op[1]=='PUTE':
                record = current.pop(ref, None)
                if record and ref not in pending and record['attempts'][-1]['status']!='failed':
                    record['pute'] = row
                    box_owners = {key: owner for key, owner in box_owners.items() if owner is not record}
            continue
        if reply and command is not None:
            add_reply(command, row, reply)
            continue
        if row.get('ref') and row.get('modbus') is not None and (references is None or row['ref'] in references):
            records[row['number']] = dict(ticket=row, attempts=[], pute=None)
            pending[row['ref']] = row
            current.pop(row['ref'], None)
            box_owners = {key: owner for key, owner in box_owners.items() if key[0]!=row['ref']}
    result = {number: _finish(record) for number, record in records.items()}
    return result


def _file(store, path, revision, cancel, references):
    cache = store._analysis_cache
    with cache.work_lock(('get_build', revision, path), lambda: store._check_cancel(cancel)):
        loaded = {ref: cache.get(('get_reference', revision, path, ref)) for ref in references}
        missing = [ref for ref, value in loaded.items() if value is None]
        if missing:
            owners = {}
            def remembered():
                for row in _targeted_rows(store, path, missing, revision, cancel):
                    if row.get('ref') and row.get('modbus') is not None:
                        owners[row['number']] = row['ref']
                    yield row
            outcomes = _analyze(store, remembered(), cancel, set(missing))
            for ref in missing:
                result = {number: value for number, value in outcomes.items() if owners.get(number)==ref}
                loaded[ref] = result
                weight = sum(1+len(d['attempts'])+len(d['responses']) for d in result.values())
                cache.put(('get_reference', revision, path, ref), result, weight)
        return {number: value for result in loaded.values() for number, value in result.items()}


def diagnostics(store, tickets, revision, cancel=None):
    store._check_cancel(cancel)
    files = {}
    scopes = {}
    for ticket in tickets:
        scopes.setdefault(ticket['path'], set()).add(ticket['ref'])
    result = {}
    for ticket in tickets:
        path = ticket['path']
        if path not in files:
            files[path] = _file(store, path, revision, cancel, scopes[path])
        result[(path, ticket['number'])] = files[path].get(ticket['number'],
            dict(status='missing', failed=False, message='GET non trouvé pour cette occurrence',
                 attempts=[], get=None, pute=None, responses=[]))
    return result

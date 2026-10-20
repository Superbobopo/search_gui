"""Exact ticket/GET/IDPK associations and bounded, publication-aware query caches."""
from collections import OrderedDict, defaultdict
from contextlib import contextmanager
from datetime import datetime
import json
import os
import threading

from trace_parser import timestamp_parts, trace_start_datetime
from trace_sources import robot_scope, path_key


class AnalysisCache:
    def __init__(self, budget=24000):
        self.lock = threading.RLock()
        self.items = OrderedDict()
        self.weight = 0
        self.budget = budget
        self.running = {}

    @contextmanager
    def work_lock(self, key, check):
        """One worker per analysis key; waiting workers remain cancellable."""
        while True:
            check()
            with self.lock:
                event = self.running.get(key)
                if event is None:
                    event = self.running[key] = threading.Event()
                    break
            event.wait(.02)
        try:
            check()
            yield
        finally:
            with self.lock:
                self.running.pop(key, None)
                event.set()

    def get(self, key):
        with self.lock:
            item = self.items.get(key)
            if item is not None:
                self.items.move_to_end(key)
                return item[0]

    def put(self, key, value, weight=1):
        weight = max(1, weight)
        with self.lock:
            if key in self.items:
                self.weight -= self.items.pop(key)[1]
            if weight <= self.budget:
                self.items[key] = value, weight
                self.weight += weight
            while self.weight > self.budget or len(self.items) > 4096:
                self.weight -= self.items.popitem(last=False)[1][1]
        return value

    def clear(self):
        with self.lock:
            self.items.clear()
            self.weight = 0


def dimensions(row):
    value = row.get('dims') or {}
    return json.loads(value) if isinstance(value, str) else value


def mismatch(first, second):
    a, b = dimensions(first), dimensions(second)
    return any(abs(a[key] - b[key]) > 2 for key in ('l', 'H', 'L') if key in a and key in b)


def dated(store, path, revision):
    key = ('dated', revision, path)
    cached = store._analysis_cache.get(key)
    if cached is not None:
        return cached
    rows = store._rows('''SELECT text FROM lines WHERE file_id=(SELECT id FROM files WHERE path=?)
                         ORDER BY number LIMIT 8''', (path,))
    return store._analysis_cache.put(key, trace_start_datetime('\n'.join(r['text'] for r in rows), None) is not None)


def timed(store, row, revision):
    day, hour = timestamp_parts(row['text'])
    if not day or not dated(store, row['path'], revision):
        return False
    try:
        base = datetime.fromtimestamp(row['time'])
        value = datetime.fromisoformat(f'{base:%Y-%m}-{int(day):02}T{hour}')
        return abs(value.timestamp() - row['time']) < .0001
    except (ValueError, OverflowError):
        return False


def metadata(store, revision):
    key = ('metadata', revision)
    result = store._analysis_cache.get(key)
    if result is None:
        result = store.metadata()
        store._analysis_cache.put(key, result, len(result))
    return result


def _identity(row):
    return row.get('ref'), row.get('sn'), row.get('lot')


def file_associations(store, path, revision, cancel=None, *, references, rows=None):
    """Only associate requested identities; never expand to every file reference."""
    refs = set(references)
    cache = store._analysis_cache
    with cache.work_lock(('association_build', revision, path), lambda: store._check_cancel(cancel)):
        result = {ref: cache.get(('association', revision, path, ref)) for ref in refs}
        missing = {ref for ref, value in result.items() if value is None}
        if not missing:
            return result
        if rows is None:
            marks = ','.join('?' for _ in missing)
            rows = store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                WHERE f.path=? AND l.ref IN (''' + marks + ') AND l.modbus IS NOT NULL UNION ALL '
                + 'SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id WHERE f.path=? '
                + 'AND l.op_ref IN (' + marks + ') AND (l.loading=1 OR l.dims IS NOT NULL)',
                (path, *missing, path, *missing), cancel)
        grouped = defaultdict(list)
        for row in rows:
            ref = row['ref'] if row['modbus'] is not None else row['op_ref']
            if ref in missing:
                grouped[ref].append(row)
        from get_diagnostics import diagnostics
        # Ordinary ticket -> GET -> IDPK histories need no gripper-response scan.
        # Only an unowned GET carrying a measurement can require retry ownership.
        retry_refs = set()
        for ref, group in grouped.items():
            group.sort(key=lambda row: row['number'])
            pending = False
            active_owned = None
            seen_owned = False
            for row in group:
                store._check_cancel(cancel)
                if row['modbus'] is not None:
                    pending = True
                elif row['loading']:
                    active_owned = pending
                    seen_owned = seen_owned or pending
                    pending = False
                elif row['dims'] and active_owned is False and seen_owned:
                    retry_refs.add(ref)
                    break
        tickets = [r for ref in retry_refs for r in grouped[ref] if r['modbus'] is not None]
        outcomes = diagnostics(store, tickets, revision, cancel) if tickets else {}
        for ref in missing:
            store._check_cancel(cancel)
            group = grouped[ref]
            tickets = [row for row in group if row['modbus'] is not None]
            associations, records = {}, []
            retry_owners = {}
            for owner in tickets:
                outcome = outcomes.get((path, owner['number']), {})
                for attempt in outcome.get('attempts', [])[1:]:
                    retry_owners[attempt['get']['number']] = owner
            owned_records = {}
            ticket = active = None
            for row in group:
                if row['modbus'] is not None:
                    ticket = row
                elif row['loading']:
                    owner = ticket if ticket is not None else retry_owners.get(row['number'])
                    active = owned_records.get(owner['number']) if owner is not None else None
                    if active is None:
                        active = {'ticket': owner, 'get': row, 'measurements': []}
                        if owner is not None:
                            owned_records[owner['number']] = active
                            associations.setdefault((path, owner['number']), []).append(active)
                            records.append(active)
                    # A later GET without a fresh ticket cannot inherit this
                    # ticket's SN/lot. It may be another box of the same SKU.
                    ticket = None
                elif row['dims'] and active is not None:
                    active['measurements'].append(row)
            payload = associations, records
            result[ref] = payload
            cache.put(('association', revision, path, ref), payload, len(group)+len(records))
        return result


def current_measurement(store, ticket, revision, cancel=None):
    payload = file_associations(store, ticket['path'], revision, cancel, references=[ticket['ref']]).get(ticket['ref'], ({}, []))
    selected = payload[0].get((ticket['path'], ticket['number']), [])
    measurements = [event for item in selected for event in item['measurements']]
    ambiguous = len(selected)>1 or len(measurements)>1
    return dict(revision=revision, status='ambiguous_measurement' if ambiguous else
                'current' if measurements else 'missing_measurement',
                current=measurements[0] if measurements and not ambiguous else None,
                message='Plusieurs mesures possibles pour cette occurrence.' if ambiguous else
                '' if measurements else 'Aucune mesure IDPK rattachable à cette occurrence.')


def _load_references(store, scope, references, revision, cancel):
    loaded = {ref: store._analysis_cache.get(('reference', revision, scope, ref)) for ref in references}
    missing = [ref for ref, payload in loaded.items() if payload is None]
    if not missing:
        return loaded
    files = [m['path'] for m in metadata(store, revision) if m['kind'] == 'omega' and robot_scope(m['path'])[0] == scope]
    for start in range(0, len(missing), 80):
        store._check_cancel(cancel)
        batch = missing[start:start + 80]
        refs = ','.join('?' for _ in batch)
        grouped = defaultdict(list)
        for file_start in range(0, len(files), 200):
            paths = files[file_start:file_start + 200]
            marks = ','.join('?' for _ in paths)
            rows = store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                WHERE f.path IN (''' + marks + ') AND l.ref IN (' + refs + ''') AND l.modbus IS NOT NULL
                UNION ALL SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                WHERE f.path IN (''' + marks + ') AND l.op_ref IN (' + refs + ''')
                AND (l.loading=1 OR l.dims IS NOT NULL)''', (*paths, *batch, *paths, *batch), cancel)
            for row in rows:
                grouped[row['ref'] if row['modbus'] is not None else row['op_ref']].append(row)
        needed = defaultdict(set)
        payloads = {ref: ({}, []) for ref in batch}
        for ref in batch:
            for row in grouped[ref]:
                needed[row['path']].add(ref)
        # A large file may exceed the cache budget. Still analyze it only once
        # per batch, rather than once per requested reference.
        for path, needed_refs in needed.items():
            input_rows = [row for ref in needed_refs for row in grouped[ref] if row['path']==path]
            local_file = file_associations(store, path, revision, cancel, references=needed_refs, rows=input_rows)
            for ref in needed_refs:
                local = local_file.get(ref, ({}, []))
                payloads[ref][0].update(local[0])
                payloads[ref][1].extend(local[1])
        for ref, payload in payloads.items():
            loaded[ref] = payload
            store._analysis_cache.put(('reference', revision, scope, ref), payload,
                                      len(grouped[ref]) + len(payload[1]))
    return loaded


def comparisons(store, tickets, revision, cancel=None):
    """Batch missing reference histories; each result belongs to one exact ticket."""
    groups = defaultdict(set)
    for row in tickets:
        if row.get('ref'):
            groups[robot_scope(row['path'])[0]].add(row['ref'])
    result = {}
    for scope, refs in groups.items():
        loaded = _load_references(store, scope, refs, revision, cancel)
        for row in tickets:
            if robot_scope(row['path'])[0] == scope:
                result[(row['path'], row['number'])] = compare(store, row, revision, cancel, loaded.get(row['ref']))
    return result


def compare(store, ticket, revision, cancel=None, payload=None):
    key = ('comparison', revision, ticket['path'], ticket['number'])
    cached = store._analysis_cache.get(key)
    if cached is not None:
        return cached
    store._check_cancel(cancel)
    result = {'status': 'missing_measurement', 'different': False, 'pair': None, 'current': None,
              'revision': revision, 'identity': _identity(ticket), 'message': 'Aucune mesure IDPK rattachable à cette occurrence.',
              'basis_label': 'Dernière mesure comparable', 'candidates': []}
    scope = robot_scope(ticket['path'])[0]
    if not ticket.get('ref'):
        return result
    ref_key = ('reference', revision, scope, ticket['ref'])
    payload = payload or store._analysis_cache.get(ref_key)
    if payload is None:
        payload = _load_references(store, scope, [ticket['ref']], revision, cancel)[ticket['ref']]
    associations, records = payload
    selected = associations.get((ticket['path'], ticket['number']), [])
    measurements = [event for item in selected for event in item['measurements']]
    if len(selected) > 1 or len(measurements) > 1:
        result.update(status='ambiguous_measurement', message='Plusieurs mesures possibles pour cette occurrence.')
    elif measurements:
        current = measurements[0]
        result['current'] = current
        if not any(key in dimensions(current) for key in ('l','H','L')):
            result.update(status='missing_dimensions', message='Aucune dimension IDPK exploitable pour cette occurrence.')
        elif not all(_identity(ticket)):
            result.update(status='incomplete_identity', message='Comparaison indisponible : Référence, SN ou CodeLot manquant.')
        elif not timed(store, current, revision):
            result.update(status='unknown_time', message='Horodatage de mesure inexploitable pour la comparaison.')
        else:
            previous, seen = [], set()
            for item in records:
                store._check_cancel(cancel)
                if _identity(item['ticket']) != _identity(ticket) or len(item['measurements']) != 1:
                    continue
                event = item['measurements'][0]
                if not any(key in dimensions(current) and key in dimensions(event) for key in ('l','H','L')):
                    continue
                earlier = (event['time'] < current['time'] or
                           (event['time'] == current['time'] and event['path'] == current['path'] and event['number'] < current['number']))
                if not earlier or not timed(store, event, revision):
                    continue
                identity = (path_key(os.path.dirname(event['path'])), event['time'], event['text'])
                if identity not in seen:
                    previous.append(event)
                    seen.add(identity)
            previous.sort(key=lambda item: (item['time'], item['number'] if item['path'] == current['path'] else -1), reverse=True)
            different = [event for event in previous if mismatch(current, event)]
            candidates = different or previous
            if candidates:
                latest = candidates[0]['time']
                closest = [event for event in candidates if event['time'] == latest]
                # A generation's line order resolves simultaneous measurements in that file.
                if all(event['path'] == closest[0]['path'] for event in closest):
                    closest = closest[:1]
                result.update(different=bool(different), status='different' if different else 'same',
                              basis_label='Mesure antérieure différente' if different else 'Dernière mesure comparable',
                              message='')
                if len(closest) > 1:
                    result.update(status='ambiguous_previous', candidates=closest,
                                  message='Plusieurs mesures antérieures au même instant : choisissez la comparaison.')
                else:
                    result['pair'] = (current, closest[0])
            else:
                result.update(status='no_previous', message='Aucune mesure antérieure comparable avec les mêmes Référence, SN et CodeLot.')
    return store._analysis_cache.put(key, result, 2)

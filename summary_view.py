"""Summary presentation, using saved identities and one outer scroll area."""

from datetime import datetime
import json

import customtkinter as ctk
from ui_components import ActionButton
from selectable_text import make_selectable
from occurrence_comparison import mismatch
from summary_widgets import SummaryWidgets, configure_changed


def dimensions(event):
    values = event.get('dims') or {}
    return json.loads(values) if isinstance(values, str) else values


def comparison_rows(pair):
    current, previous = map(dimensions, pair)
    return [(key, current.get(key), previous.get(key),
             current[key] - previous[key] if key in current and key in previous else None)
            for key in ('l', 'H', 'L')]


def comparison_mismatch(pair):
    # Same tolerance and shared dimensions as TraceStore.differs / comparison popup.
    return mismatch(*pair)


def dimension_text(event):
    values = dimensions(event)
    return '   '.join(f'{key} = {values.get(key, "—")}' for key in ('l', 'H', 'L'))


def render_summary(app, data):
    container = getattr(app, 'summary_frame', None)
    if container is None or not container.winfo_exists():
        return
    palette = getattr(app, '_palette', {})
    comparison = (data or {}).get('_summary_comparison', (data or {}).get('_summary_current', {}))
    pair = comparison.get('pair') or ()
    base_key = (id(data), id((data or {}).get('locations')), len((data or {}).get('locations', ())),
                (data or {}).get('location_lot'), getattr(app, '_sources_verified', False),
                getattr(app, '_indexed_complete', False), repr((data or {}).get('get_diagnostic')),
                repr((data or {}).get('row')), repr((data or {}).get('path')),
                repr((data or {}).get('get')), repr((data or {}).get('pute')))
    key = (id(data), id((data or {}).get('locations')), len((data or {}).get('locations', ())),
           (data or {}).get('location_lot'),
           tuple((str(event['path']), getattr(event['path'], 'version_id', None), event['number'],
                  repr(event.get('dims'))) for event in pair), comparison.get('error'),
           comparison.get('status'), comparison.get('different'), comparison.get('message'),
           '_summary_comparison' in (data or {}), '_summary_current' in (data or {}),
           comparison.get('revision'), repr(comparison.get('current')),
           getattr(app, '_sources_verified', False),
           getattr(app, '_indexed_complete', False), base_key, tuple(palette.items()))
    cached = getattr(container, '_summary_cache', None)
    # A hidden page cancels its workers. Returning to unchanged widgets must resume
    # the unfinished analysis even when no visual update is necessary.
    if (cached is not None and cached[0] is data and data and '_summary_comparison' not in data
            and app._sources_verified and app._indexed_complete
            and getattr(getattr(app, 'selection', None), 'details', None) is data):
        app._load_summary_comparison(data)
    if cached is not None and cached[0] is data and cached[1] == key:
        return
    if cached is not None and cached[0] is data and cached[1][:-1] == key[:-1]:
        from ui_theme import apply_widget_theme
        apply_widget_theme(container, palette)
        container._summary_cache = (data, key)
        return
    partial = bool(data and cached is not None and cached[0] is data
                   and getattr(container, '_summary_base_key', None) == base_key
                   and getattr(container, '_summary_idpk_frame', None) is not None
                   and container._summary_idpk_frame.winfo_exists())
    if partial and cached[1][-1] != key[-1]:
        from ui_theme import apply_widget_theme
        apply_widget_theme(container, palette)
    container._summary_base_key = base_key
    # Retain data itself, so a recycled Python id cannot identify another selection.
    container._summary_cache = (data, key)
    pool = getattr(container, '_summary_widgets', None)
    if pool is None:
        pool = container._summary_widgets = SummaryWidgets(app, container)
    pool.begin(data, container._summary_idpk_frame if partial else None)
    if not partial:
        app.summary_targets = {}
    else:
        for name in ('idpk', 'previous_idpk'):
            app.summary_targets.pop(name, None)

    foreground = palette.get('text', '#172a3a')
    surface = palette.get('surface', '#ffffff')

    bind_open = pool.bind_open

    def label(parent, text, *, bold=False, centered=False, small=False, color=None, target=None):
        widget = pool.label(parent, text=text, text_color=color or foreground,
                              anchor='center' if centered else 'w',
                              justify='center' if centered else 'left', wraplength=600,
                              font=('Segoe UI', 11 if small else 12, 'bold' if bold else 'normal'))
        pool.pack(widget, fill='x', padx=12, pady=4)
        pool.bind_resize(parent)
        make_selectable(widget)
        if target:
            bind_open(widget, target)
        return widget

    def card(key, title=None):
        frame = pool.frame(container, key=key, fg_color=surface, corner_radius=8,
                              border_width=1, border_color=palette.get('border', '#d6dfe8'))
        pool.pack(frame, fill='x', padx=12, pady=4)
        if title:
            label(frame, title, bold=True, centered=True)
        return frame

    header = None if partial else pool.frame(container, key='identity', fg_color='transparent')
    if not data:
        card('empty', 'Aucune occurrence sélectionnée.')
        pool.finish()
        return
    if not partial:
        row = data['row']
        pool.pack(header, fill='x', padx=12, pady=(12, 8))
        for column, (name, value) in enumerate((('Référence', row.get('ref')), ('ID', row.get('modbus')),
                                               ('SN', row.get('sn')), ('CodeLot', row.get('lot')))):
            header.grid_columnconfigure(column, weight=1, uniform='identity')
            cell = pool.frame(header, fg_color='transparent')
            pool.grid(cell, row=0, column=column, sticky='nsew', padx=4)
            label(cell, name, centered=True, small=True)
            field = label(cell, value or '—', centered=True, bold=True)
            if name == 'Référence' and value:
                configure_changed(field, text_color=('#0f766e', '#35d6c6'))
                pool.bind_reference(field)


    for name in (('idpk',) if partial else ('get', 'pute', 'idpk')):
        event = data.get(name)
        if partial:
            frame = container._summary_idpk_frame
            pool.used.add(frame)
        else:
            frame = card(name)
        if name == 'idpk':
            container._summary_idpk_frame = frame
        if name == 'idpk':
            result = data.get('_summary_comparison', data.get('_summary_current', {}))
            pair = result.get('pair') if app._sources_verified and app._indexed_complete else None
            current = result.get('current') if 'status' in result else pair[0] if pair else None
            previous = pair[1] if pair else None
            mismatch = result.get('different', comparison_mismatch(pair) if pair else False) if '_summary_comparison' in data and app._sources_verified and app._indexed_complete else False
            measurements = pool.frame(frame, fg_color='transparent')
            pool.pack(measurements, fill='x', padx=8, pady=3)
            for column, (title, measurement, key) in enumerate((('IDPK actuelle', current, 'idpk'),
                                                                (result.get('basis_label', 'Précédente'), previous, 'previous_idpk'))):
                measurements.grid_columnconfigure(column, weight=1, uniform='idpk')
                cell = pool.frame(measurements, fg_color='transparent')
                pool.grid(cell, row=0, column=column, sticky='nsew', padx=4)
                line = pool.frame(cell, fg_color='transparent')
                pool.pack(line, anchor='center', pady=2)
                title_label = pool.label(line, text=title + ' : ', text_color=foreground,
                                           font=('Segoe UI', 11))
                pool.pack(title_label, side='left')
                make_selectable(title_label)
                values = pool.label(line, text=dimension_text(measurement) if measurement else '—',
                                       text_color=('#b42318', '#ff9b91') if column == 0 and mismatch else foreground,
                                       font=('Segoe UI', 11, 'bold'))
                pool.pack(values, side='left')
                make_selectable(values)
                if measurement:
                    app.summary_targets[key] = values
                    bind_open(cell, measurement)
                    bind_open(line, measurement)
                    for child in line.winfo_children():
                        if isinstance(child, ctk.CTkLabel):
                            bind_open(child, measurement)
                    when = datetime.fromtimestamp(measurement['time'])
                    label(cell, f'{when:%d/%m/%Y %H:%M:%S} · '
                          + app._compact_trace_label(measurement['path']), centered=True,
                          small=True, target=measurement)
            if not app._sources_verified or not app._indexed_complete:
                label(frame, 'Comparaison disponible après vérification des sources.', centered=True, small=True)
            elif '_summary_comparison' not in data:
                label(frame, 'Recherche de la mesure antérieure…' if '_summary_current' in data
                      else 'Recherche de la mesure actuelle…', centered=True, small=True)
                app._load_summary_comparison(data)
            else:
                if result.get('error'):
                    label(frame, 'Comparaison indisponible : ' + result['error'], centered=True, small=True)
                elif result.get('message') or not pair:
                    label(frame, result.get('message') or 'Aucune comparaison disponible.', centered=True, small=True)
        elif event:
            bind_open(frame, event)
            diagnostic = data.get('get_diagnostic', {}) if name=='get' else {}
            verified = app._sources_verified and app._indexed_complete
            failed = verified and diagnostic.get('failed')
            attempts = diagnostic.get('attempts', []) if verified else []
            if failed and len(attempts) > 1:
                for i, attempt in enumerate(attempts):
                    success = attempt['status'] == 'success'
                    role = 'positive' if success else 'warning'
                    block = pool.frame(frame, key=('attempt', i), disposable=i>=4, fg_color=palette[role + '_bg'], corner_radius=6)
                    pool.pack(block, fill='x', padx=6, pady=6)
                    outcome = ('GET réussi' if success else 'Échec GET' if attempt['status']=='failed'
                               else 'Résultat GET inconnu')
                    heading = label(block, f'Tentative {i + 1} — {outcome}', bold=True,
                                    color=palette[role + '_text'])
                    app.summary_targets[f'get_attempt_status_{i}'] = heading
                    if i == 0:
                        app.summary_targets['get_status'] = heading
                    command = label(block, attempt['get']['text'], target=attempt['get'],
                                    color=('#00589b', '#8ac7ff') if success else palette['warning_text'])
                    app.summary_targets['get' if i == 0 else f'get_attempt_{i}'] = command
                continue
            if failed:
                configure_changed(frame, fg_color=palette['warning_bg'])
            app.summary_targets[name] = label(frame, event['text'], target=event,
                                               color=palette['warning_text'] if failed else ('#00589b', '#8ac7ff') if name == 'get'
                                               else ('#9c5d00', '#ffd18a'))
            if name=='get' and not (verified and diagnostic.get('status')=='success'):
                app.summary_targets['get_status'] = label(frame,
                    diagnostic.get('message', 'Résultat GET inconnu') if verified else 'Diagnostic GET disponible après vérification des sources.',
                    bold=True, color=palette['warning_text'] if failed else foreground)
        else:
            label(frame, f'{name.upper()} : non trouvé pour cette occurrence', centered=True, small=True)

    if partial:
        pool.finish()
        return

    locations = data.get('locations', [])
    location_lot = data.get('location_lot')
    title = (f'Derniers emplacements connus · CodeLot {location_lot} ({len(locations)})'
             if location_lot is not None else f'Derniers emplacements connus par lot ({len(locations)})')
    group = card('locations', title)
    if not locations:
        label(group, 'Aucun emplacement connu pour ce lot.' if location_lot is not None
              else 'Aucun emplacement connu pour cette référence.', centered=True)
        more = getattr(group, '_summary_more', None)
        if more is not None:
            more.pack_forget()
        pool.finish()
        return
    items = pool.frame(group, fg_color='transparent')
    pool.pack(items, fill='x', padx=8, pady=4)
    more = getattr(group, '_summary_more', None)
    if more is None:
        more = group._summary_more = ActionButton(group, text='Afficher plus')
    from ui_theme import apply_widget_theme
    apply_widget_theme(more, palette)
    pool.pack(more, pady=(4, 10))
    locations_key = (data['path'], data['row'].get('number'), repr(locations), location_lot)
    shown = 0
    limit = getattr(group, '_summary_shown', 20) if getattr(group, '_summary_locations_key', None)==locations_key else 20
    group._summary_locations_key = locations_key

    def append_locations(initial=False):
        nonlocal shown
        end = min(limit if initial else shown + 20, len(locations))
        for index in range(shown, end):
            event = locations[index]
            item = pool.frame(items, key=('location', index), disposable=index>=20, fg_color='transparent', border_width=1,
                                 border_color=palette.get('border', '#d6dfe8'))
            pool.pack(item, fill='x', pady=4)
            when = datetime.fromtimestamp(event['time'])
            bind_open(item, event)
            label(item, f"Lot {event['lot_hint']} · {when:%d/%m/%Y %H:%M:%S}", bold=True, target=event)
            label(item, app._compact_trace_label(event['path']), small=True, target=event)
            target = label(item, event['text'], small=True, target=event)
            app.summary_targets.setdefault('location', target)
        shown = group._summary_shown = end
        if shown == len(locations):
            more.pack_forget()
        else:
            more.configure(text=f'Afficher plus · {shown}/{len(locations)} lots')
    more.configure(command=lambda: append_locations() if pool.can_open() else None)
    append_locations(initial=True)
    pool.finish()

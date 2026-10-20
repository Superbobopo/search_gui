"""Summary presentation, using saved identities and one outer scroll area."""

from datetime import datetime
import json

import customtkinter as ctk
from ui_components import ActionButton, font as ui_font
from selectable_text import make_selectable
from occurrence_comparison import mismatch


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
    if app.result_mode.get() != 'Synthèse':
        return
    palette = getattr(app, '_palette', {})
    comparison = (data or {}).get('_summary_comparison', {})
    pair = comparison.get('pair') or ()
    key = (id(data), id((data or {}).get('locations')), len((data or {}).get('locations', ())),
           (data or {}).get('location_lot'),
           tuple((str(event['path']), getattr(event['path'], 'version_id', None), event['number'],
                  repr(event.get('dims'))) for event in pair), comparison.get('error'),
           comparison.get('status'), comparison.get('different'), comparison.get('message'),
           '_summary_comparison' in (data or {}), getattr(app, '_sources_verified', False),
           getattr(app, '_indexed_complete', False), tuple(palette.items()))
    cached = getattr(container, '_summary_cache', None)
    if cached is not None and cached[0] is data and cached[1] == key:
        return
    if cached is not None and cached[0] is data and cached[1][:-1] == key[:-1]:
        from ui_theme import apply_widget_theme
        apply_widget_theme(container, palette)
        container._summary_cache = (data, key)
        return
    # Retain data itself, so a recycled Python id cannot identify another selection.
    container._summary_cache = (data, key)
    for child in container.winfo_children():
        child.destroy()
    app.summary_targets = {}
    foreground = palette.get('text', '#172a3a')
    surface = palette.get('surface', '#ffffff')

    def bind_open(widget, event):
        widget.configure(cursor='hand2')
        def open_trace(_event):
            app._open_indexed_file(event['path'], event['number'], event['text'])
            return 'break'
        widget.bind('<Double-Button-1>', open_trace, add='+')
        if hasattr(widget, 'selection_text'):
            widget.selection_text.bind('<Double-Button-1>', open_trace, add='+')

    def label(parent, text, *, bold=False, centered=False, small=False, color=None, target=None):
        widget = ctk.CTkLabel(parent, text=text, text_color=color or foreground,
                              anchor='center' if centered else 'w',
                              justify='center' if centered else 'left', wraplength=600,
                              font=('Segoe UI', 11 if small else 12, 'bold' if bold else 'normal'))
        widget.pack(fill='x', padx=12, pady=4)
        def resize(event, w=widget):
            width = max(80, event.width - 32)
            if w.winfo_exists() and w.cget('wraplength') != width:
                w.configure(wraplength=width)
        parent.bind('<Configure>', resize, add='+')
        make_selectable(widget)
        if target:
            bind_open(widget, target)
        return widget

    def card(title=None):
        frame = ctk.CTkFrame(container, fg_color=surface, corner_radius=8,
                              border_width=1, border_color=palette.get('border', '#d6dfe8'))
        frame.pack(fill='x', padx=12, pady=4)
        if title:
            label(frame, title, bold=True, centered=True)
        return frame

    if not data:
        card('Aucune occurrence sélectionnée.')
        return
    row = data['row']
    header = ctk.CTkFrame(container, fg_color='transparent')
    header.pack(fill='x', padx=12, pady=(12, 8))
    for column, (name, value) in enumerate((('Référence', row.get('ref')), ('ID', row.get('modbus')),
                                           ('SN', row.get('sn')), ('CodeLot', row.get('lot')))):
        header.grid_columnconfigure(column, weight=1, uniform='identity')
        cell = ctk.CTkFrame(header, fg_color='transparent')
        cell.grid(row=0, column=column, sticky='nsew', padx=4)
        label(cell, name, centered=True, small=True)
        field = label(cell, value or '—', centered=True, bold=True)
        if name == 'Référence' and value:
            field.configure(text_color=('#0f766e', '#35d6c6'))
            field.sync_selection_text()
            # Activate on release, so dragging still selects text for copying.
            gesture = {}
            def press(event, state=gesture):
                state.update(x=event.x, y=event.y, dragged=False)
            def drag(event, state=gesture):
                if abs(event.x - state.get('x', event.x)) > 4 or abs(event.y - state.get('y', event.y)) > 4:
                    state['dragged'] = True
            def release(event, state=gesture):
                moved = (abs(event.x - state.get('x', event.x)) > 4
                         or abs(event.y - state.get('y', event.y)) > 4)
                if (not moved and not state.get('dragged', False) and app.modbus.get()
                        and app.selection is not None and app.selection.details is data):
                    app.on_click_refprod(event)
                    return 'break'
            field.selection_text.bind('<ButtonPress-1>', press, add='+')
            field.selection_text.bind('<B1-Motion>', drag, add='+')
            field.selection_text.bind('<ButtonRelease-1>', release, add='+')

    for name in ('get', 'pute', 'idpk'):
        event = data.get(name)
        frame = card()
        if name == 'idpk':
            result = data.get('_summary_comparison', {})
            pair = result.get('pair') if app._sources_verified and app._indexed_complete else None
            current = result.get('current') if 'status' in result else pair[0] if pair else None
            previous = pair[1] if pair else None
            mismatch = result.get('different', comparison_mismatch(pair) if pair else False) if app._sources_verified and app._indexed_complete else False
            measurements = ctk.CTkFrame(frame, fg_color='transparent')
            measurements.pack(fill='x', padx=8, pady=3)
            for column, (title, measurement, key) in enumerate((('IDPK actuelle', current, 'idpk'),
                                                                (result.get('basis_label', 'Précédente'), previous, 'previous_idpk'))):
                measurements.grid_columnconfigure(column, weight=1, uniform='idpk')
                cell = ctk.CTkFrame(measurements, fg_color='transparent')
                cell.grid(row=0, column=column, sticky='nsew', padx=4)
                line = ctk.CTkFrame(cell, fg_color='transparent')
                line.pack(anchor='center', pady=2)
                title_label = ctk.CTkLabel(line, text=title + ' : ', text_color=foreground,
                                           font=('Segoe UI', 11))
                title_label.pack(side='left')
                make_selectable(title_label)
                values = ctk.CTkLabel(line, text=dimension_text(measurement) if measurement else '—',
                                       text_color=('#b42318', '#ff9b91') if column == 0 and mismatch else foreground,
                                       font=('Segoe UI', 11, 'bold'))
                values.pack(side='left')
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
                label(frame, 'Recherche de la mesure précédente…', centered=True, small=True)
                app._load_summary_comparison(data)
            else:
                if result.get('error'):
                    label(frame, 'Comparaison indisponible : ' + result['error'], centered=True, small=True)
                elif result.get('message') or not pair:
                    label(frame, result.get('message') or 'Aucune comparaison disponible.', centered=True, small=True)
        elif event:
            bind_open(frame, event)
            app.summary_targets[name] = label(frame, event['text'], target=event,
                                               color=('#00589b', '#8ac7ff') if name == 'get'
                                               else ('#9c5d00', '#ffd18a'))
        else:
            label(frame, f'{name.upper()} : non trouvé pour cette occurrence', centered=True, small=True)

    locations = data.get('locations', [])
    location_lot = data.get('location_lot')
    title = (f'Derniers emplacements connus · CodeLot {location_lot} ({len(locations)})'
             if location_lot is not None else f'Derniers emplacements connus par lot ({len(locations)})')
    group = card(title)
    if not locations:
        label(group, 'Aucun emplacement connu pour ce lot.' if location_lot is not None
              else 'Aucun emplacement connu pour cette référence.', centered=True)
        return
    items = ctk.CTkFrame(group, fg_color='transparent')
    items.pack(fill='x', padx=8, pady=4)
    more = ActionButton(group, text='Afficher plus')
    from ui_theme import apply_widget_theme
    apply_widget_theme(more, palette)
    more.pack(pady=(4, 10))
    shown = 0

    def append_locations():
        nonlocal shown
        end = min(shown + 20, len(locations))
        for event in locations[shown:end]:
            item = ctk.CTkFrame(items, fg_color='transparent', border_width=1,
                                 border_color=palette.get('border', '#d6dfe8'))
            item.pack(fill='x', pady=4)
            when = datetime.fromtimestamp(event['time'])
            bind_open(item, event)
            label(item, f"Lot {event['lot_hint']} · {when:%d/%m/%Y %H:%M:%S}", bold=True, target=event)
            label(item, app._compact_trace_label(event['path']), small=True, target=event)
            target = label(item, event['text'], small=True, target=event)
            app.summary_targets.setdefault('location', target)
        shown = end
        if shown == len(locations):
            more.pack_forget()
        else:
            more.configure(text=f'Afficher plus · {shown}/{len(locations)} lots')
    more.configure(command=append_locations)
    append_locations()

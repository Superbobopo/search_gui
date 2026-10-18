"""Real Tk event regression checks, using disposable traces and SQLite only.

--script launches the actual `python search_gui.py` command in a child process.
A temporary sitecustomize installs the same test driver before its mainloop.
--desktop places both robot fixtures under a temporary Desktop/traces/R1,R2.
"""
from contextlib import ExitStack
from pathlib import Path
import os
import subprocess
import sys
import tempfile
import threading
import time
import weakref
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import indexed_ui
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from trace_store import TraceStore


def prepare():
    stack = ExitStack()
    base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
    fixtures = PROJECT / 'tests' / 'fixtures'
    def ticket(mb, ref, marker):
        return (f'  14|07:00:00.0 RefProduit={ref}<Cr>ModBusId={mb}<Cr>\n'
                f'G 14|07:00:01.0 GRIPPER GET 1x boîte(s): Code:{ref} $0101 {marker}\n'
                f'G 14|07:00:02.0 GRIPPER PUTE 1x boîte(s): Code:{ref} {marker}\n')
    paths = {}
    for robot, fixture in [('trace r1', 'current_r1.txt'), ('trace r2', 'current_r2.txt')]:
        directory = (base / 'Desktop' / 'traces' / robot[-2:].upper()
                     if '--desktop' in sys.argv else base / 'traces' / robot)
        directory.mkdir(parents=True)
        text = read_trace(fixtures / fixture)
        if robot == 'trace r1':
            text += '\n' + ticket('1194', 'OTHER_R1', 'R1_ONLY')
        else:
            for mb in range(1195, 1203):
                text += '\n' + ticket(str(mb), f'REF{mb}', f'R2_{mb}')
            text += '\n' + ticket('1194', 'REUSED', 'REUSED_SECOND')
            text += '  14|08:00:00.0 RefProduit=MISSING_GET<Cr>ModBusId=8888<Cr>\n'
        paths[robot] = str(directory / 'TracOMEGA.txt')
        Path(paths[robot]).write_bytes(text.encode('cp1252'))
        values = ['0'] * 121
        values[0] = '14|06:27:47'; values[2] = '1194'; values[6] = '1194'; values[-1] = ''
        (directory / 'TracMODBUS.txt').write_text(';'.join(values) + '\n', encoding='utf8')
    old = read_trace(fixtures / 'current_r2.txt').replace('14/10/2026', '13/10/2026').replace('14|', '13|').replace('l=48 H=21 L=47', 'l=55 H=21 L=47')
    (Path(paths['trace r2']).parent / 'TracOmega029.old').write_bytes(old.encode('cp1252'))
    stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
    stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'Desktop')))
    stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'absent')))
    stack.enter_context(patch.object(indexed_ui, 'TraceStore', side_effect=lambda: TraceStore(base / 'index.db')))
    # Freeze automatic sync after initial publication to test a modified file
    # between index refreshes. Production continues to poll every two seconds.
    def index_once(app):
        store = indexed_ui.TraceStore()
        app._index_messages.put(('ready', store, [], False))
        result = store.sync(sources.discover_trace_files())
        app._index_messages.put(('complete', result, store.metadata()))
    stack.enter_context(patch.object(indexed_ui.IndexedAppMixin, '_index_loop', index_once))
    return stack, paths


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def exercise(app, paths, script=False):
    failures, alerts = [], []
    previous_unraisable_hook = sys.unraisablehook
    sys.unraisablehook = lambda event: failures.append(
        (event.exc_type, event.exc_value, event.exc_traceback))
    app.report_callback_exception = lambda *error: failures.append(error)
    app.geometry('1100x900+18000+18000')
    def pump(condition=lambda: True, seconds=5):
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            app.update()
            for window in windows():
                if window.winfo_x() < 10000:
                    window.geometry('+18000+18000')
            if failures:
                raise failures[0][1]
            if condition():
                return
            time.sleep(.005)
        raise AssertionError('Operation timed out: ' + app.status.get())
    def windows():
        return [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)]
    def titled(prefix):
        return [w for w in windows() if w.title().startswith(prefix)]
    def search(value, reference=False):
        app.keyword.delete(0, 'end'); app.keyword.insert(0, value)
        app.modbus.set(not reference); app.reference.set(reference)
        app.on_reference_check() if reference else app.on_modbus_check()
        app.keyword._entry.focus_force()
        app.keyword._entry.event_generate('<Return>')
        pump(lambda: app.selection is not None and 'Recherche…' not in app.status.get())
    def choose(path, reference=None):
        for _ in range(len(app._groups)):
            if app.selection.path == path:
                break
            old = app.selection
            app.btn_next_file.invoke()
            pump(lambda: app.selection is not None and app.selection is not old)
        assert app.selection.path == path
        if reference:
            for _ in range(len(app.lines)):
                if app.selection.reference == reference:
                    break
                app.on_down(); pump(lambda: app.selection is not None)
            assert app.selection.reference == reference
    stamp = 1000
    def click_tag(tag, double=False):
        nonlocal stamp
        start = app.txt.tag_ranges(tag)[0]
        pos = f'{start} +1lines +5chars' if double else f'{start} +5chars'
        app.txt.see(pos); pump()
        box = app.txt.bbox(pos)
        assert box, (tag, pos)
        x, y = box[0] + 2, box[1] + 2
        app.txt.event_generate('<Motion>', x=x, y=y)
        for i in range(2 if double else 1):
            app.txt.event_generate('<ButtonPress-1>', x=x, y=y, time=stamp + i * 100)
            app.txt.event_generate('<ButtonRelease-1>', x=x, y=y, time=stamp + i * 100 + 20)
        stamp += 1000
    def open_event(key):
        event, path = app.selection.details[key], app.selection.path
        previous = set(titled('Vue de'))
        click_tag('clickable_' + key, double=True)
        pump(lambda: bool(set(titled('Vue de')) - previous))
        window = (set(titled('Vue de')) - previous).pop()
        text = next(w for w in descendants(window) if isinstance(w, gui.tk.Text))
        pump(lambda: bool(text.tag_ranges('highlight')))
        begin, end = text.tag_ranges('highlight')
        assert text.get(begin, end) == event['text']
        assert int(str(begin).split('.')[0]) == event['number'] + 1
        assert text.get('1.0', 'end-1c').splitlines() == read_trace(path).splitlines()
        assert sources.trace_label(path) in window.title()
        return window, text
    def switch_result(value, keyboard=False):
        switch = next(w for w in descendants(app)
                      if isinstance(w, gui.ctk.CTkSegmentedButton)
                      and value in w._buttons_dict)
        button = switch._buttons_dict[value]
        if keyboard:
            button._canvas.focus_force()
            button._canvas.event_generate('<Return>')
        else:
            button.invoke()
        pump(lambda: app.result_mode.get() == value)
    def check_label_contrast(label):
        native = label._text_label if isinstance(label, gui.ctk.CTkButton) else label._label
        foreground, background = native.cget('fg'), native.cget('bg')
        desired = label._apply_appearance_mode(label.cget('text_color'))
        assert app.winfo_rgb(foreground) == app.winfo_rgb(desired), (label.cget('text'), foreground, desired)
        def luminance(color):
            channels = [value / 65535 for value in app.winfo_rgb(color)]
            linear = [value / 12.92 if value <= .04045 else ((value + .055) / 1.055) ** 2.4
                      for value in channels]
            return sum(value * weight for value, weight in zip(linear, (.2126, .7152, .0722)))
        low, high = sorted((luminance(foreground), luminance(background)))
        ratio = (high + .05) / (low + .05)
        assert ratio >= 4.5, (label.cget('text'), foreground, background, ratio)
    def check_design_states():
        # Inspect rendered native labels after real Tk hover events, including
        # controls created after the theme was selected.
        app.btn_gear.invoke(); pump()
        check_label_contrast(app.sn_label)
        check_label_contrast(app.codelot_label)
        app.btn_gear.invoke(); pump()
        for mode in ('Classique', 'Synthèse'):
            previous_boxes = [w for w in descendants(app.summary_frame)
                              if isinstance(w, gui.ctk.CTkTextbox)]
            pending = {handle for box in previous_boxes
                       for handle in box._pending_callbacks}
            switch_result(mode)
            assert all(not box._pending_callbacks for box in previous_boxes)
            assert not pending.intersection(app.tk.splitlist(app.tk.call('after', 'info')))
            for button in (app.btn_gear, app.btn_copy_path,
                           *app.result_switch._buttons_dict.values()):
                check_label_contrast(button)
                button._canvas.event_generate('<Enter>'); pump()
                check_label_contrast(button)
                button._canvas.event_generate('<Leave>'); pump()
        for name, button in (('search', app.btn_search_menu),
                             ('tracking', app.btn_tracking_robot)):
            icon = app._nav_icons[name]
            for surface in (button._canvas, icon):
                surface.event_generate('<Enter>'); pump()
                assert app.winfo_rgb(icon.cget('bg')) == app.winfo_rgb(button._text_label.cget('bg'))
                surface.event_generate('<Leave>'); pump()
                assert app.winfo_rgb(icon.cget('bg')) == app.winfo_rgb(button._text_label.cget('bg'))
    def check_presentation():
        selected = app.selection
        original_text = app.txt.get('1.0', 'end-1c')
        original_styles = {tag: tuple(app.txt.tag_cget(tag, key) for key in
                                    ('background', 'foreground', 'justify', 'font'))
                           for tag in ('date', 'occ', 'getcard', 'putecard', 'refcard')}
        original_index_status = app.index_status.get()
        request_status = app.status.get()
        app._index_messages.put(('progress', 1, 2, selected.path, False, []))
        pump(lambda: 'Actualisation 1/2' in app.index_status.get())
        assert app.status.get() == request_status
        app.index_status.set(original_index_status)
        switch_result('Synthèse')
        assert app.selection is selected
        assert set(app.summary_buttons) == {'get', 'idpk', 'pute'}
        summary_text = ' '.join(str(w.cget('text')) for w in descendants(app.summary_frame)
                                if isinstance(w, gui.ctk.CTkLabel))
        assert selected.reference in summary_text and 'l = 48' in summary_text
        for value in ('Sombre', 'Clair', 'Système'):
            app.theme_menu._dropdown_callback(value)
            pump(lambda: app.theme_mode.get() == value)
            assert app.selection is selected
            for parent in (app.sidebar, app.summary_frame):
                for label in descendants(parent):
                    if isinstance(label, gui.ctk.CTkLabel):
                        check_label_contrast(label)
            check_label_contrast(app.keyword_label)
            check_design_states()
            assert app.txt.get('1.0', 'end-1c') == original_text
            assert {tag: tuple(app.txt.tag_cget(tag, key) for key in
                               ('background', 'foreground', 'justify', 'font'))
                    for tag in original_styles} == original_styles
            for key in ('get', 'idpk', 'pute'):
                event = selected.details[key]
                previous = set(titled('Vue de'))
                with patch.object(app, '_render_indexed_file_view', wraps=app._render_indexed_file_view) as render:
                    app.summary_buttons[key].invoke()
                    pump(lambda: bool(set(titled('Vue de')) - previous))
                    assert render.call_args.args[0] == event.get('path', selected.path)
                    assert render.call_args.args[2] == event['number']
                window = (set(titled('Vue de')) - previous).pop()
                text = next(w for w in descendants(window) if isinstance(w, gui.tk.Text))
                pump(lambda: bool(text.tag_ranges('highlight')))
                begin, end = text.tag_ranges('highlight')
                assert text.get(begin, end) == event['text']
                assert int(str(begin).split('.')[0]) == event['number'] + 1
                assert text.get('1.0', 'end-1c').splitlines() == read_trace(selected.path).splitlines()
                assert sources.trace_label(selected.path) in window.title()
                assert window.cget('bg') == app._palette['bg']
                if value == 'Sombre':
                    app.theme_menu._dropdown_callback('Clair')
                    pump(lambda: window.cget('bg') == app._palette['bg'])
                    app.theme_menu._dropdown_callback('Sombre')
                    pump(lambda: window.cget('bg') == app._palette['bg'])
                status_reference = weakref.ref(window.trace_controller.status)
                window.destroy()
                assert window.trace_controller.status is None
                assert status_reference() is None
        switch_result('Classique', keyboard=True)
        switch_result('Synthèse', keyboard=True)
        switch_result('Classique', keyboard=True)
        assert app.selection is selected
        assert app.txt.get('1.0', 'end-1c') == original_text
        app.btn_gear.invoke(); pump()
        try:
            app.geometry('860x650+18000+18000')
            for scale in (1, 1.5, 2):
                gui.ctk.set_widget_scaling(scale)
                gui.ctk.set_window_scaling(scale)
                pump()
                controls = (app.keyword, app.btn_search, app.btn_gear, app.sn_entry,
                            app.codelot_entry, app.btn_copy_path, app.result_switch,
                            app.btn_up, app.btn_down, app.btn_compare_idpk,
                            app.btn_adjacent, app.btn_open_trace, app.btn_erreurs)
                for control in controls:
                    assert control.winfo_ismapped()
                    assert control.winfo_rootx() >= app.workspace.winfo_rootx()
                    assert control.winfo_rootx() + control.winfo_width() <= app.winfo_rootx() + app.winfo_width()
                for icon in app._nav_icons.values():
                    assert icon.winfo_width() == round(22 * scale)
                    assert icon.winfo_x() == round(10 * scale)
        finally:
            gui.ctk.set_widget_scaling(1)
            gui.ctk.set_window_scaling(1)
            app.geometry('1100x900+18000+18000')
            app.btn_gear.invoke(); pump()
        print('OK filtres crees en theme sombre, survols, icones et commandes a 100/150/200 %', flush=True)
        print('OK themes Systeme/Clair/Sombre, contrastes reels, styles Classique conserves, Synthese clavier et boutons GET/IDPK/PUTE exacts', flush=True)
    def check_large_local_search():
        original_lines = read_trace(app.selection.path).splitlines()
        lines = original_lines + [f'  14|12:00:00.0 common sample line {number:06d}' for number in range(30000)]
        target_number = len(original_lines) + 20000
        lines[target_number] += ' UNIQUE_TARGET'
        previous = set(titled('Vue de'))
        with patch.object(app.store, 'file_lines', return_value=lines):
            click_tag('clickable_get', double=True)
            pump(lambda: bool(set(titled('Vue de')) - previous))
        window = (set(titled('Vue de')) - previous).pop()
        controller = window.trace_controller
        text, local = controller.text, controller.search_entry
        ticks = []
        heartbeat_active = [True]
        def beat():
            if heartbeat_active[0]:
                ticks.append(time.monotonic())
                app.after(5, beat)
        beat()
        try:
            pump(lambda: controller.loaded, seconds=15)
            assert text.get('1.0', 'end-1c').splitlines() == lines
            local.focus_force()
            def type_filter(value):
                local.delete(0, 'end'); local.insert(0, value)
                local.event_generate('<KeyRelease>', keysym='a')
            completed = controller.search_completed
            generations = controller.search_generation
            for value in ('OBSOLETE_LOCAL_FILTER', 'common', 'UNIQUE_TARGET'):
                type_filter(value)
                pump()
            pump(lambda: controller.search_ready, seconds=10)
            assert controller.search_generation == generations + 3
            assert controller.search_completed == completed + 1
            expected_position = f"{target_number + 1}.{lines[target_number].index('UNIQUE_TARGET')}"
            assert controller.search_state['positions'] == [expected_position]
            local.event_generate('<Return>')
            pump(lambda: bool(text.tag_ranges('mainline')))
            assert int(text.index('insert').split('.')[0]) == target_number + 1
            type_filter('common')
            pump(lambda: controller.search_ready, seconds=10)
            buttons = [w for w in descendants(window) if isinstance(w, gui.ttk.Button)]
            buttons[0].invoke()
            assert text.index('insert') == controller.search_state['positions'][-1]
            buttons[1].invoke()
            assert text.index('insert') == controller.search_state['positions'][0]
            started_ticks = len(ticks)
            type_filter('common')
            pump(lambda: bool(controller.search_state['positions']) and not controller.search_ready, seconds=10)
            type_filter('UNIQUE_TARGET')
            pump(lambda: controller.search_ready, seconds=10)
            assert len(controller.search_state['positions']) == 1
            assert len(ticks) > started_ticks + 2
            type_filter('common')
            pump(lambda: bool(controller.search_state['positions']) and not controller.search_ready, seconds=10)
            window.destroy()
            pump(lambda: controller.closed)
            deadline = time.monotonic() + .25
            while time.monotonic() < deadline:
                pump(); time.sleep(.005)
            assert not controller._pending and not failures
            assert str(text) not in app._text_tokens
        finally:
            heartbeat_active[0] = False
            if window.winfo_exists():
                window.destroy()
        print('OK grande trace 30000 lignes: frappe temporisee, recherche annulee, navigation, heartbeat et fermeture sans callback obsolete', flush=True)
    def current_label():
        return ' '.join(str(w.cget('text')) for w in app.current_frame.winfo_children())
    with patch.object(indexed_ui.messagebox, 'showinfo', side_effect=lambda *args, **kw: alerts.append(args)), \
         patch.object(indexed_ui.messagebox, 'showwarning', side_effect=lambda *args, **kw: alerts.append(args)):
        try:
            pump(lambda: app._indexed_complete)
            if script:
                assert app.__class__.__module__ == '__main__'
                assert app.__class__ is not gui.SearchApp
            search('1194'); choose(paths['trace r2'], 'A0061574AC')
            original = app.selection
            app.btn_adjacent.invoke(); pump(lambda: app.adj_popup is not None)
            pump(lambda: 'A0061574AC' in current_label())
            assert len(app.next_frame.winfo_children()) == 5
            print('OK bouton IDs adjacents, homonymes R1/R2, cinq voisins', flush=True)
            click_tag('clickable_refprod'); pump(lambda: app.selection is not None and app.reference.get())
            assert (app.selection.path, app.selection.number) == (original.path, original.number)
            click_tag('clickable_refprod'); pump(lambda: app.selection is not None and app.modbus.get())
            assert app.keyword.get() == '1194'
            assert (app.selection.path, app.selection.number) == (original.path, original.number)
            print('OK vrais clics ID -> Reference -> ID, occurrence conservee', flush=True)
            for key in ('get', 'pute'):
                window, text = open_event(key)
                window.destroy()
            print('OK vrais doubles-clics GET/PUTE -> contenu integral et ligne exacte', flush=True)
            assert app.txt.tag_cget('date', 'justify') == 'center'
            assert app.txt.tag_cget('occ', 'justify') == 'center'
            for tag, color in [('getcard', '#e3f2fd'), ('putecard', '#fff4e3'), ('refcard', '#e9fdf4')]:
                assert app.txt.tag_cget(tag, 'background') == color
            check_presentation()
            click_tag('clickable_refprod'); pump(lambda: app.selection is not None and app.reference.get())
            app.btn_compare_idpk.invoke()
            pump(lambda: bool(titled('Comparaison IDPK')))
            comparison = ' '.join(str(w.cget('text')) for w in descendants(titled('Comparaison IDPK')[0]) if isinstance(w, gui.tk.Label))
            assert 'l = 48' in comparison and 'l = 55' in comparison
            # A real neighboring label event updates both main selection and popup.
            app.next_frame.winfo_children()[0].event_generate('<Button-1>')
            pump(lambda: app.selection is not None and app.selection.modbus == '1195')
            pump(lambda: '1195' in current_label())
            assert app.selection.path == paths['trace r2']
            window, text = open_event('get')
            buttons = [w for w in descendants(window) if isinstance(w, gui.ttk.Button)]
            get_next = buttons[3]; pute_next = buttons[5]
            get_next.invoke(); first = text.index('insert')
            get_next.invoke(); second = text.index('insert')
            assert first != second and 'GRIPPER GET' in text.get('insert linestart', 'insert lineend')
            pute_next.invoke(); assert 'GRIPPER PUTE' in text.get('insert linestart', 'insert lineend')
            local = next(w for w in descendants(window) if isinstance(w, gui.ttk.Entry))
            local.insert(0, 'ref1195'); local.focus_force(); local.event_generate('<KeyRelease>', keysym='a')
            pump(lambda: bool(text.tag_ranges('search')))
            buttons[1].invoke(); assert text.tag_ranges('mainline')
            assert 'ref1195' in text.get('insert linestart', 'insert lineend').lower()
            window.destroy()
            print('OK clic voisin, popup actualisee, fleches GET/PUTE et recherche locale', flush=True)
            check_large_local_search()
            app.btn_erreurs.invoke(); pump(lambda: bool(titled('Vérification erreurs Robot')))
            error_label = next(w for w in descendants(titled('Vérification erreurs Robot')[0])
                               if isinstance(w, gui.tk.Label) and 'MISSING_GET' in str(w.cget('text')))
            error_label.event_generate('<Button-1>')
            pump(lambda: app.selection is not None and app.selection.modbus == '8888')
            pump(lambda: '8888' in current_label())
            assert app.selection.path == paths['trace r2']
            # Select the second use of the same ID through its own reference.
            search('REUSED', reference=True)
            pump(lambda: 'REUSED' in current_label())
            assert 'REUSED_SECOND' in app.selection.details['get']['text']
            print('OK clic erreur Robot et ID reutilise ancre sur la seconde occurrence', flush=True)
            # A slow obsolete adjacent result must not repopulate the popup.
            started, release, finished = threading.Event(), threading.Event(), threading.Event()
            original_adjacent = app.store.adjacent
            def slow_adjacent(*args):
                started.set(); release.wait(3)
                try: return original_adjacent(*args)
                finally: finished.set()
            with patch.object(app.store, 'adjacent', slow_adjacent):
                app.btn_adjacent.invoke(); pump(started.is_set)
                app.keyword.delete(0, 'end'); app.keyword.insert(0, 'ABSENT_999')
                app.perform_search(); pump(lambda: 'Aucun fichier' in app.txt.get('1.0', 'end'))
                release.set(); pump(finished.is_set); pump()
            assert app.selection is None and app._display_data is None
            assert app.btn_adjacent.cget('state') == 'disabled'
            assert not app.prev_frame.winfo_children() and not app.next_frame.winfo_children()
            assert 'Aucune occurrence' in current_label()
            search('1194'); choose(paths['trace r2'], 'A0061574AC')
            # A slow file opening belonging to a previous selection is discarded.
            started.clear(); release.clear(); finished.clear()
            original_lines = app.store.file_lines
            def slow_lines(*args):
                started.set(); release.wait(3)
                try: return original_lines(*args)
                finally: finished.set()
            previous = set(titled('Vue de'))
            with patch.object(app.store, 'file_lines', slow_lines):
                click_tag('clickable_get', double=True); pump(started.is_set)
                search('1195')
                release.set(); pump(finished.is_set)
                pump(lambda: app.tasks.results.empty())
            assert set(titled('Vue de')) == previous
            search('1194'); choose(paths['trace r2'], 'A0061574AC')
            # Detect a shifted line even before automatic reindexing.
            path = Path(app.selection.path); saved = path.read_bytes()
            path.write_bytes(b'inserted\n' + saved)
            event = app.selection.details['get']; previous = set(titled('Vue de'))
            click_tag('clickable_get', double=True)
            pump(lambda: bool(set(titled('Vue de')) - previous))
            view = (set(titled('Vue de')) - previous).pop()
            text = next(w for w in descendants(view) if isinstance(w, gui.tk.Text))
            pump(lambda: bool(text.tag_ranges('highlight')))
            assert str(text.tag_ranges('highlight')[0]) == f"{event['number'] + 2}.0"
            view.destroy()
            path.write_bytes(saved.replace(event['text'].encode('cp1252'), b'removed event'))
            click_tag('clickable_get', double=True); pump(lambda: bool(alerts))
            assert alerts[-1][0] == 'Trace modifiée'
            path.write_bytes(saved); alerts.clear()
            path.unlink()
            click_tag('clickable_get', double=True); pump(lambda: bool(alerts))
            assert 'supprimé' in alerts[-1][1]
            path.write_bytes(saved); alerts.clear()
            print('OK absence, reponses obsoletes ignorees, fichier et ligne deplaces/supprimes', flush=True)
            # Remaining tools: inspect their displayed results, not just their creation.
            app.convoyage_action('1194'); pump(lambda: bool(getattr(app, 'convoyage_window', None)))
            def labels(window):
                return ' '.join(str(w.cget('text')) for w in descendants(window) if isinstance(w, gui.tk.Label))
            pump(lambda: '1194' in labels(app.convoyage_window))
            origin = 'Bureau' if '--desktop' in sys.argv else 'Local'
            assert f'{origin} R1' in labels(app.convoyage_window) and f'{origin} R2' in labels(app.convoyage_window)
            card = next(w for w in descendants(app.convoyage_window)
                        if isinstance(w, gui.tk.Label) and '1194' in str(w.cget('text')) and w.bind('<Button-1>'))
            card.event_generate('<Button-1>')
            pump(lambda: 'Parcours détaillé du ModBusId' in labels(app.convoyage_window))
            assert 'Présent' in labels(app.convoyage_window)
            app.show_convoyage_details('1194'); app.convoyage_errors_action()
            pump(lambda: bool(titled('Parcours détaillé')) and bool(titled('Erreurs Convoyage')))
            assert 'Présent' in labels(titled('Parcours détaillé')[0])
            app.show_tracking_robot_panel(); pump(); app.tracking_keyword.insert(0, 'A0061574AC')
            app.tracking_keyword._entry.focus_force(); app.tracking_keyword._entry.event_generate('<Return>')
            pump(lambda: '14/10/26' in app.tracking_txt.get('1.0', 'end'))
            app.show_search_panel(); app.keyword.delete(0, 'end'); app.keyword.insert(0, 'A0061574AC')
            app.btn_gear.invoke(); pump(); app.sn_entry.insert(0, '343272'); app.codelot_entry.insert(0, '261013.127')
            app.sn_entry._entry.focus_force(); app.sn_entry._entry.event_generate('<Return>')
            pump(lambda: app.selection is not None and len(app._groups) == 2)
            app.btn_next_file.invoke(); pump(lambda: app.selection is not None and '029.old' in app.selection.path)
            app.btn_prev_file.invoke(); pump(lambda: app.selection is not None and 'TracOMEGA.txt' in app.selection.path)
            # Actual rapid requests, and a view change while a request is pending.
            for value in ('1194', 'NO_RESULT', '1195'):
                app.keyword.delete(0, 'end'); app.keyword.insert(0, value); app.modbus.set(True); app.reference.set(False); app.perform_search()
            pump(lambda: app.selection is not None and app.selection.modbus == '1195')
            app.keyword.delete(0, 'end'); app.keyword.insert(0, 'ABSENT_999'); app.perform_search()
            app.show_tracking_robot_panel(); pump()
            assert not alerts and not failures
            print('OK IDPK, convoyage, tracking, recherche avancee, navigation et recherches rapides', flush=True)
            with tempfile.TemporaryDirectory() as empty_folder:
                app.show_search_panel()
                app.store = TraceStore(Path(empty_folder) / 'empty.db')
                app._install_metadata([])
                app.keyword.delete(0, 'end'); app.keyword.insert(0, '1194'); app.perform_search()
                pump(lambda: 'Aucun fichier' in app.txt.get('1.0', 'end'))
                assert app.selection is None and not app._groups and not app.lines
                app.open_file_view()
                assert alerts[-1][1] == 'Aucun fichier de trace disponible.'
                alerts.clear()
            assert not failures
            print('OK aucune trace disponible: absence, etat vide et ouverture protegee', flush=True)
            print('INTERACTIONS OK (' + ('python search_gui.py' if script else 'import') + ')', flush=True)
        finally:
            app._close_indexed_app()
            app._index_thread.join(timeout=5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
            sys.unraisablehook = previous_unraisable_hook


def script_bootstrap():
    stack, paths = prepare()
    def mainloop(app, *args, **kwargs):
        try:
            exercise(app, paths, script=True)
        finally:
            stack.close()
    gui.ctk.CTk.mainloop = mainloop


def main():
    if '--script' in sys.argv:
        with tempfile.TemporaryDirectory() as folder:
            Path(folder, 'sitecustomize.py').write_text(
                'import sys\nsys.path.insert(0, ' + repr(str(PROJECT)) + ')\n'
                'try:\n'
                '    from tools.check_interactions import script_bootstrap\n'
                '    script_bootstrap()\n'
                'except BaseException:\n'
                '    import os, traceback\n'
                '    traceback.print_exc()\n'
                '    os._exit(2)\n', encoding='utf8')
            env = dict(os.environ, PYTHONPATH=folder)
            command = [sys.executable, str(PROJECT / 'search_gui.py')]
            if '--desktop' in sys.argv:
                command.append('--desktop')
            result = subprocess.run(command, env=env, cwd=folder, timeout=60,
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            print(result.stdout, end='')
            raise SystemExit(result.returncode)
    stack, paths = prepare()
    with stack:
        exercise(gui.SearchApp(), paths)


if __name__ == '__main__':
    from tools.ui_test_support import run_gui_check
    run_gui_check(main)

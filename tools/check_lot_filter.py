"""Actual filter controls and location cards on a private desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
import search_gui as gui
import trace_sources as sources
from tests.test_lot_locations import REFERENCE, LOT, trace_block
from tools.ui_test_support import run_gui_check


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        for robot, text in (('R1', trace_block(LOT) + trace_block(LOT + '0', 7, '1195')),
                            ('R2', trace_block(LOT, 8, '1196') + trace_block('OTHER', 9, '1197')
                             + '\n14|10:00:00.0 RefProduit=' + REFERENCE +
                             '<Cr>CodeLot=NO_LOCATION<Cr>SN=NONE<Cr>ModBusId=8888<Cr>\n')):
            path = base / 'traces' / robot / 'TracOmega019.old'
            path.parent.mkdir(parents=True)
            path.write_text(text, encoding='utf8')
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)

        def pump(condition=lambda: True):
            deadline = time.monotonic() + 8
            while time.monotonic() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError(app.status.get())

        def search(lot):
            previous = app.selection
            app.codelot_entry.delete(0, 'end')
            app.codelot_entry.insert(0, lot)
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and app.selection is not previous
                 and app.selection.details.get('location_lot') == (lot or None))

        def labels():
            return [str(w.cget('text')) for w in descendants(app.summary_frame) if isinstance(w, ctk.CTkLabel)]

        def assert_filtered(lot):
            assert app.selection.details['row']['lot'] == lot
            assert {event['lot_hint'] for event in app.selection.details['locations']} <= {lot}
            displayed = [text.split(' · ', 1)[0] for text in labels() if text.startswith('Lot ')]
            assert displayed == ['Lot ' + lot], displayed

        try:
            pump(lambda: app._sources_verified)
            app.mode_radios['ref'].invoke()
            app.keyword.insert(0, REFERENCE)
            app.btn_gear.invoke()
            search(LOT)
            assert_filtered(LOT)
            event = app.selection.details['locations'][0]
            assert '08:28:07.6' in event['text']  # later OTHER must not hide this one
            for _ in range(2):
                prior = app.selection
                app.btn_next_file.invoke()
                pump(lambda: app.selection is not None and app.selection is not prior)
                assert_filtered(LOT)
            app.result_switch._buttons_dict['Classique'].invoke()
            text = app.txt.get('1.0', 'end')
            assert 'Lot ' + LOT + ' —' in text and 'Lot ' + LOT + '0 —' not in text and 'Lot OTHER' not in text
            app.result_switch._buttons_dict['Synthèse'].invoke()
            app._change_theme('Sombre')
            assert_filtered(LOT)
            app.btn_tracking_robot.invoke()
            app.btn_search_menu.invoke()
            assert_filtered(LOT)
            if not app.advanced_open:
                app.btn_gear.invoke()
            # Clear, prefix-collision and no-location searches must invalidate the display cache.
            search(LOT + '0')
            assert_filtered(LOT + '0')
            search('NO_LOCATION')
            assert app.selection.details['locations'] == []
            assert 'Aucun emplacement connu pour ce lot.' in labels()
            search('')
            assert {event['lot_hint'] for event in app.selection.details['locations']} == {LOT + '0', 'OTHER'}
            print('OK CodeLot exact : vrais Filtres/Rechercher, occurrences et emplacements filtres '
                  'avant choix du dernier PUTE, prefixe distinct, navigation R1/R2, Classique/Synthese, '
                  'theme/cache de panneaux, absence sans repli et effacement du filtre')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)

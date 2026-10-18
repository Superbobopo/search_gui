"""A trace window with cooperative local search and prepared event navigation.

All widget access, including the scheduled search slices, stays on the Tk thread.
The original line numbers remain the coordinate system for every operation.
"""

from bisect import bisect_left, bisect_right
import time
import tkinter as tk
from tkinter import scrolledtext, ttk

from app_assets import apply_icon
from trace_sources import trace_label


class TraceViewController:
    DEBOUNCE_MS = 200
    SEARCH_LINES = 400
    SLICE_SECONDS = .006
    MATCH_BATCH = 150

    def __init__(self, window, text, search_entry, status):
        self.window = window
        self.text = text
        self.search_entry = search_entry
        self.status = status
        self.loaded = False
        self.closed = False
        self.search_ready = False
        self.search_generation = 0
        self.search_completed = 0
        self.search_state = {'positions': [], 'current': -1}
        self.event_positions = {'get': [], 'pute': []}
        self.buttons = {}
        self._pending = set()
        self._debounce = None
        self._cursor = '1.0'
        self._term = ''
        self._pending_navigation = None
        search_entry.bind('<KeyRelease>', self.schedule_search)
        search_entry.bind('<Return>', self._enter_search)
        window.bind('<Destroy>', self._on_destroy, add='+')

    def _schedule(self, delay, callback):
        if self.closed:
            return None
        handle = None
        def run():
            self._pending.discard(handle)
            if not self.closed:
                callback()
        handle = self.window.after(delay, run)
        self._pending.add(handle)
        return handle

    def _cancel_pending(self):
        for handle in tuple(self._pending):
            try:
                self.window.after_cancel(handle)
            except tk.TclError:
                pass
        self._pending.clear()
        self._debounce = None

    def _on_destroy(self, event):
        if event.widget is self.window:
            self.closed = True
            self.search_generation += 1
            self._cancel_pending()
            # The window/controller cycle can otherwise keep this Tk variable
            # alive until cyclic GC happens in a data worker. Release it now,
            # while Tcl can still be accessed from the UI thread.
            self.status = None

    def prepared_lines(self, lines):
        for number, line in enumerate(lines):
            if 'GRIPPER GET' in line:
                self.event_positions['get'].append(number)
            if 'GRIPPER PUTE' in line:
                self.event_positions['pute'].append(number)
            yield line + '\n'

    def finish_loading(self, highlight_line, count):
        if self.closed:
            return
        self.loaded = True
        for key in ('get', 'pute'):
            for button in self.buttons.get(key, ()):
                button.configure(state='normal' if self.event_positions[key] else 'disabled')
        if highlight_line is not None and 0 <= highlight_line < count:
            start = f'{highlight_line + 1}.0'
            self.text.tag_add('highlight', start, f'{start} lineend')
            self.text.see(start)
            self.text.mark_set('insert', start)
        self.status.set(f'{count:,} lignes'.replace(',', ' '))
        if self.search_entry.get().strip():
            self.schedule_search()

    def schedule_search(self, event=None):
        # A Return event has its own binding; releasing it must not cancel navigation.
        if event is not None and event.keysym in ('Return', 'KP_Enter'):
            return
        self.search_generation += 1
        self._cancel_pending()
        self.search_ready = False
        self.search_state['positions'].clear()
        self.search_state['current'] = -1
        self._pending_navigation = None
        self.text.tag_remove('search', '1.0', 'end')
        if not self.loaded:
            self.status.set('Chargement de la trace…')
            return
        if not self.search_entry.get().strip():
            self.search_ready = True
            self.status.set('Aucun filtre local')
            return
        self.status.set('Recherche locale…')
        generation = self.search_generation
        self._debounce = self._schedule(self.DEBOUNCE_MS, lambda: self._start_search(generation))

    def _start_search(self, generation):
        if generation != self.search_generation or not self.loaded:
            return
        self._debounce = None
        self._term = self.search_entry.get().strip()
        self._cursor = '1.0'
        self._search_slice(generation)

    def _search_slice(self, generation):
        if self.closed or generation != self.search_generation:
            return
        deadline = time.perf_counter() + self.SLICE_SECONDS
        ranges = []
        end = self.text.index('end-1c')
        while time.perf_counter() < deadline and len(ranges) < self.MATCH_BATCH * 2:
            boundary = self.text.index(f'{self._cursor} +{self.SEARCH_LINES}lines linestart')
            if self.text.compare(boundary, '>', end):
                boundary = end
            found = self.text.search(self._term, self._cursor, nocase=True, stopindex=boundary)
            if found:
                finish = self.text.index(f'{found}+{len(self._term)}c')
                self.search_state['positions'].append(found)
                ranges.extend((found, finish))
                self._cursor = finish
            else:
                self._cursor = boundary
                if self.text.compare(boundary, '>=', end):
                    if ranges:
                        self.text.tag_add('search', *ranges)
                    self.search_ready = True
                    self.search_completed += 1
                    self.status.set(f"{len(self.search_state['positions'])} correspondance(s)")
                    if self._pending_navigation is not None:
                        offset, self._pending_navigation = self._pending_navigation, None
                        self.goto_search(offset)
                    return
        if ranges:
            self.text.tag_add('search', *ranges)
        self._schedule(1, lambda: self._search_slice(generation))

    def _enter_search(self, event=None):
        self.goto_search(1)
        return 'break'

    def goto_search(self, offset):
        if not self.search_ready:
            self._pending_navigation = offset
            return
        positions = self.search_state['positions']
        if not positions:
            return
        current = self.search_state['current']
        self.search_state['current'] = ((0 if offset > 0 else len(positions) - 1) if current < 0
                                        else (current + offset) % len(positions))
        self._show_position(positions[self.search_state['current']], 'mainline')

    def navigate(self, key, offset):
        positions = self.event_positions[key]
        if not self.loaded or not positions:
            return
        current = int(self.text.index('insert').split('.')[0]) - 1
        target = bisect_right(positions, current) if offset > 0 else bisect_left(positions, current) - 1
        number = positions[target % len(positions)]
        self._show_position(f'{number + 1}.0', 'getline' if key == 'get' else 'puteline')

    def _show_position(self, position, tag):
        self.text.see(position)
        self.text.mark_set('insert', position)
        self.text.tag_remove(tag, '1.0', 'end')
        self.text.tag_add(tag, position, f'{position} lineend')


def show_trace_view(app, path, lines, highlight_line=None):
    palette = getattr(app, '_palette', {})
    bg = palette.get('bg', '#f7f9fa')
    surface = palette.get('surface', '#fafbfc')
    foreground = palette.get('text', '#1e293b')
    top = tk.Toplevel(app)
    apply_icon(top)
    top.title(f'Vue de {trace_label(path)}')
    top.geometry('980x650')
    top.configure(bg=bg)
    top.columnconfigure(0, weight=1)
    top.rowconfigure(1, weight=1)
    toolbar = tk.Frame(top, bg=bg)
    toolbar.grid(row=0, column=0, sticky='ew', padx=16, pady=(10, 4))
    toolbar.columnconfigure(1, weight=1)
    tk.Label(toolbar, text='Recherche dans ce fichier :', bg=bg, fg=foreground,
             font=('Segoe UI', 10)).grid(row=0, column=0, padx=(0, 8))
    local = ttk.Entry(toolbar, font=('Segoe UI', 11), width=25)
    local.grid(row=0, column=1, sticky='ew')
    navigation = tk.Frame(toolbar, bg=bg)
    navigation.grid(row=0, column=2, padx=(8, 0))
    previous = ttk.Button(navigation, text='↑', width=3)
    previous.pack(side='left')
    following = ttk.Button(navigation, text='↓', width=3)
    following.pack(side='left', padx=(2, 0))
    events = tk.Frame(toolbar, bg=bg)
    events.grid(row=1, column=0, columnspan=3, sticky='w', pady=(8, 2))
    event_buttons = {}
    for key, title in [('get', 'GET'), ('pute', 'PUTE')]:
        group = tk.Frame(events, bg=bg)
        group.pack(side='left', padx=(0, 16))
        tk.Label(group, text=title, font=('Segoe UI', 10, 'bold'), bg=bg,
                 fg=foreground).pack(side='left', padx=(0, 6))
        before = ttk.Button(group, text='⟨', width=3, state='disabled')
        before.pack(side='left')
        after = ttk.Button(group, text='⟩', width=3, state='disabled')
        after.pack(side='left', padx=(2, 0))
        event_buttons[key] = (before, after)
    text = scrolledtext.ScrolledText(top, wrap='word', font=('Consolas', 11),
                                   bg=surface, fg=foreground, insertbackground=foreground,
                                   borderwidth=0, padx=8, pady=8)
    text.grid(row=1, column=0, sticky='nsew', padx=12, pady=6)
    for tag, color in [('mainline', '#e3f2fd'), ('getline', '#b3e5fc'),
                       ('puteline', '#ffcdd2'), ('search', '#fff59d'), ('highlight', '#ffe66d')]:
        text.tag_configure(tag, background=color, foreground='#172a3a')
    text.tag_configure('mainline', font=('Consolas', 11, 'bold'))
    status = tk.StringVar(value='Chargement de la trace…')
    tk.Label(top, textvariable=status, bg=bg, fg=palette.get('muted', '#405060'),
             anchor='w', font=('Segoe UI', 10)).grid(row=2, column=0, sticky='ew', padx=16, pady=(2, 10))
    controller = TraceViewController(top, text, local, status)
    controller.buttons = event_buttons
    top.trace_controller = controller
    previous.configure(command=lambda: controller.goto_search(-1))
    following.configure(command=lambda: controller.goto_search(1))
    for key, (before, after) in event_buttons.items():
        before.configure(command=lambda kind=key: controller.navigate(kind, -1))
        after.configure(command=lambda kind=key: controller.navigate(kind, 1))
    if hasattr(app, '_theme_window'):
        app._theme_window(top)
    app._insert_text_batches(text, controller.prepared_lines(lines),
                             lambda: controller.finish_loading(highlight_line, len(lines)))
    top.lift()
    return top

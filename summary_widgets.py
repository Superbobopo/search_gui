"""One bounded set of reusable summary widgets, owned by the Tk thread."""
import tkinter as tk
import customtkinter as ctk


def configure_changed(widget, **options):
    changed = {key:value for key,value in options.items() if widget.cget(key)!=value}
    if changed:
        # Theme adapters must remember the current business color, not a former role.
        originals = getattr(widget, '_trace_theme_colors', {})
        for key in changed:
            originals.pop(key, None)
        widget.configure(**changed)
        if hasattr(widget, 'sync_selection_text'):
            widget.sync_selection_text()


class SummaryWidgets:
    def __init__(self, app, container):
        self.app, self.container = app, container
        self.widgets = {}
        self.data = None
        self.loading = False
        self.cleanup_timer = None
        container.bind('<Destroy>', self._destroyed, add='+')

    def _destroyed(self, event):
        if event.widget is self.container and self.cleanup_timer is not None:
            try:
                self.container.after_cancel(self.cleanup_timer)
            except tk.TclError:
                pass
            self.cleanup_timer = None

    def begin(self, data, scope=None):
        self.data, self.loading = data, False
        self.scope = str(scope or self.container) + '.'
        self.used, self.counters = set(), {}

    def _widget(self, kind, parent, key, options, disposable=False):
        if key is None:
            counter = (parent, kind)
            key = self.counters.get(counter, 0)
            self.counters[counter] = key+1
        identity = (parent, kind, key)
        widget = self.widgets.get(identity)
        if widget is None:
            widget = (ctk.CTkFrame if kind=='frame' else ctk.CTkLabel)(parent, **options)
            widget._summary_disposable = disposable
            self.widgets[identity] = widget
        else:
            if kind=='label':
                options.pop('wraplength', None)
            configure_changed(widget, **options)
        self.used.add(widget)
        widget._summary_active = True
        widget._summary_event = None
        return widget

    def frame(self, parent, key=None, *, disposable=False, **options):
        return self._widget('frame', parent, key, options, disposable)

    def label(self, parent, key=None, **options):
        return self._widget('label', parent, key, options)

    def pack(self, widget, **options):
        layout = ('pack', options)
        if getattr(widget, '_summary_layout', None)!=layout or widget.winfo_manager()!='pack':
            widget.pack(**options)
            widget._summary_layout = layout

    def grid(self, widget, **options):
        layout = ('grid', options)
        if getattr(widget, '_summary_layout', None)!=layout or widget.winfo_manager()!='grid':
            widget.grid(**options)
            widget._summary_layout = layout

    def can_open(self):
        selection = getattr(self.app, 'selection', None)
        return not self.loading and selection is not None and selection.details is self.data

    def bind_open(self, widget, event):
        widget._summary_event = event
        if getattr(widget, '_summary_open_bound', False):
            return
        widget._summary_open_bound = True
        configure_changed(widget, cursor='hand2')

        def open_trace(_event):
            target = widget._summary_event
            if self.can_open() and target:
                self.app._open_indexed_file(target['path'], target['number'], target['text'])
            return 'break'

        widget.bind('<Double-Button-1>', open_trace, add='+')
        if hasattr(widget, 'selection_text'):
            widget.selection_text.bind('<Double-Button-1>', open_trace, add='+')

    def bind_reference(self, field):
        if getattr(field, '_summary_reference_bound', False):
            return
        field._summary_reference_bound = True
        gesture = {}

        def press(event):
            gesture.update(x=event.x, y=event.y, dragged=False)

        def moved(event):
            return abs(event.x-gesture.get('x', event.x))>4 or abs(event.y-gesture.get('y', event.y))>4

        def drag(event):
            gesture['dragged'] = gesture.get('dragged', False) or moved(event)

        def release(event):
            if not moved(event) and not gesture.get('dragged') and self.app.modbus.get() and self.can_open():
                self.app.on_click_refprod(event)
                return 'break'

        for sequence, callback in (('<ButtonPress-1>', press), ('<B1-Motion>', drag), ('<ButtonRelease-1>', release)):
            field.selection_text.bind(sequence, callback, add='+')

    def bind_resize(self, parent):
        if getattr(parent, '_summary_resize_bound', False):
            return
        parent._summary_resize_bound = True

        def resize(event):
            width = max(80, event.width-32)
            for (owner, kind, _), widget in tuple(self.widgets.items()):
                if owner is parent and kind=='label' and getattr(widget, '_summary_active', False):
                    configure_changed(widget, wraplength=width)

        parent.bind('<Configure>', resize, add='+')

    def finish(self):
        for widget in self.widgets.values():
            if str(widget).startswith(self.scope) and widget not in self.used:
                widget._summary_active = False
                widget._summary_event = None
                manager = widget.winfo_manager()
                if manager=='pack':
                    widget.pack_forget()
                elif manager=='grid':
                    widget.grid_remove()
                if isinstance(widget, ctk.CTkLabel):
                    configure_changed(widget, text='')
        if self.cleanup_timer is None and any(getattr(w, '_summary_disposable', False)
                                              and not w._summary_active for w in self.widgets.values()):
            self.cleanup_timer = self.container.after(25, self._cleanup)

    def _cleanup(self):
        self.cleanup_timer = None
        # Release one excess card per turn, never all expanded lots in one callback.
        victim = next((w for w in self.widgets.values() if getattr(w, '_summary_disposable', False)
                       and not w._summary_active), None)
        if victim is None:
            return
        prefix = str(victim)+'.'
        self.widgets = {key:w for key,w in self.widgets.items() if w is not victim and not str(w).startswith(prefix)}
        victim.destroy()
        self.cleanup_timer = self.container.after(25, self._cleanup)


def set_summary_loading(app):
    container = getattr(app, 'summary_frame', None)
    pool = getattr(container, '_summary_widgets', None)
    if pool is not None:
        pool.loading = True
    if hasattr(app, 'status'):
        app.status.set('Chargement…')


def reuse_summary_analysis(previous, data, revision):
    """Copy only verified analysis for the exact, unchanged occurrence and inputs."""
    if not previous or not data or previous.get('_summary_revision')!=revision:
        return
    old_path, new_path = previous['path'], data['path']
    if (str(old_path), getattr(old_path, 'version_id', None), previous['row']['number']) != (
            str(new_path), getattr(new_path, 'version_id', None), data['row']['number']):
        return
    if any(previous.get(key)!=data.get(key) for key in ('row','get','pute','get_diagnostic','locations','location_lot')):
        return
    for key in ('_summary_current', '_summary_comparison'):
        result = previous.get(key)
        if result is not None and result.get('revision')==revision and not result.get('error'):
            data[key] = result

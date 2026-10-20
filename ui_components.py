"""Shared visual components; display models never replace trace identities."""
import math
import tkinter as tk
import weakref

import customtkinter as ctk

from selectable_text import make_selectable
from ui_theme import LIGHT, DARK, palette_for

SIZES = dict(title=20, section=14, body=12, meta=11)


def font(role='body', bold=False, native=False):
    size = SIZES[role]
    if native and tk._default_root is not None:
        try:
            size = round(size*ctk.ScalingTracker.get_widget_scaling(tk._default_root))
        except (KeyError, AttributeError):
            pass
    return ('Segoe UI', -size if native else size, 'bold' if bold else 'normal')


def selectable(parent, **options):
    options.setdefault('font', font())
    label = ctk.CTkLabel(parent, anchor='w', justify='left', height=0, **options)
    make_selectable(label)
    return label


class ActionButton(ctk.CTkButton):
    """Three explicit roles, identical metrics across all application commands."""
    def __init__(self, parent, *, role='secondary', text_size=None, **options):
        self.ui_role, self._button_ready = role, False
        for option in ('fg_color', 'hover_color', 'text_color', 'text_color_disabled',
                       'border_color', 'border_width', 'font', 'height', 'corner_radius', 'style'):
            options.pop(option, None)
        if 0<options.get('width', 100)<32:
            options['width'] = max(32, options['width']*10 if options['width']<25 else 32)
        super().__init__(parent, height=32, corner_radius=6, border_width=0,
                         font=('Segoe UI', text_size or SIZES['body']), **options)
        self._button_ready = True
        self.apply_theme()

    def configure(self, **options):
        if not self._button_ready:
            return super().configure(**options)
        style = options.pop('style', None)
        role = options.pop('role', None)
        if role is not None or style is not None:
            self.ui_role = role or ('primary' if 'Active' in style else 'secondary')
        recolor = role is not None or style is not None
        if 'state' in options:
            if options['state']==self.cget('state'):
                options.pop('state')
            else:
                recolor = True
        for option in ('fg_color', 'hover_color', 'text_color', 'text_color_disabled', 'border_color'):
            recolor = options.pop(option, None) is not None or recolor
        # Preserve the common metrics even when an old theme adapter supplies them.
        for option in ('height', 'corner_radius', 'font', 'border_width'):
            options.pop(option, None)
        if options:
            super().configure(**options)
        if recolor:
            self.apply_theme()

    def apply_theme(self, colors=None):
        colors = colors or palette_for(ctk.get_appearance_mode())
        background, foreground, hover = {
            'primary': ('accent', 'accent_text', 'hover'),
            'quiet': (None, 'muted', 'raised'),
        }.get(self.ui_role, ('raised', 'text', 'border'))
        if self.cget('state')=='disabled':
            background, foreground, hover = 'raised', 'muted', 'raised'
        values = dict(fg_color=colors[background] if background else 'transparent',
                      text_color=colors[foreground], hover_color=colors[hover], text_color_disabled=colors['muted'])
        if values!=getattr(self, '_button_colors', None):
            self._button_colors = values
            super().configure(**values)


def register_view(app, view):
    if not hasattr(app, '_ui_views'):
        app._ui_views = weakref.WeakSet()
    app._ui_views.add(view)


def present_popup(popup, parent, width=1180, height=740, close=None, *, transient=True):
    from app_assets import apply_icon
    from window_placement import show_centered_popup, monitor_work_area, native_window_bounds
    import os
    apply_icon(popup)
    popup.bind('<Escape>', lambda event: (close or popup.destroy)(), add='+')
    show_centered_popup(popup, parent, width, height, transient=transient)
    work = monitor_work_area(parent)
    extra_width = extra_height = 0
    if os.name=='nt':
        _, bounds = native_window_bounds(popup)
        extra_width = bounds[2]-bounds[0]-popup.winfo_width()
        extra_height = bounds[3]-bounds[1]-popup.winfo_height()
    popup.maxsize(min(width, work[2]-work[0]-extra_width), min(height, work[3]-work[1]-extra_height))


def style_labels(labels, colors, role='text', background='surface'):
    for label in labels:
        label.configure(text_color=colors[role], bg_color=colors[background])
        if hasattr(label, 'sync_selection_text'):
            label.sync_selection_text()


class WindowHeader(ctk.CTkFrame):
    def __init__(self, parent, title, subtitle='', close=None):
        super().__init__(parent, corner_radius=0, fg_color='transparent')
        self.grid_columnconfigure(0, weight=1)
        self.title = ctk.CTkLabel(self, text=title, anchor='w', font=font('title', True))
        self.title.grid(row=0, column=0, sticky='ew')
        self.subtitle = ctk.CTkLabel(self, text=subtitle, anchor='w', font=font('meta'))
        self.subtitle.grid(row=1, column=0, sticky='ew', pady=(2, 0))
        self.close_button = ActionButton(self, text='Fermer', width=80, command=close)
        if close:
            self.close_button.grid(row=0, column=1, rowspan=2, padx=(12, 0))

    def apply_theme(self, colors):
        self.configure(fg_color=colors['bg'], bg_color=colors['bg'])
        self.title.configure(text_color=colors['text'])
        self.subtitle.configure(text_color=colors['muted'])
        self.close_button.apply_theme(colors)


class WindowShell(ctk.CTkFrame):
    def __init__(self, parent, app, title, subtitle=''):
        super().__init__(parent, corner_radius=0)
        self.app, self._theme_deferred = app, True
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.header = WindowHeader(self, title, subtitle, parent.destroy)
        self.header.grid(row=0, column=0, sticky='ew', padx=20, pady=(16, 14))
        self.title, self.close_button = self.header.title, self.header.close_button
        self.body = ctk.CTkFrame(self, corner_radius=12)
        self.body.grid(row=1, column=0, sticky='nsew', padx=20, pady=(0, 20))
        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(0, weight=1)
        register_view(app, self)
        WindowShell.apply_theme(self)

    def apply_theme(self):
        colors = self.app._palette
        self.configure(fg_color=colors['bg'])
        self.body.configure(fg_color=colors['surface'], bg_color=colors['bg'])
        self.header.apply_theme(colors)


class SplitView(ctk.CTkFrame):
    def __init__(self, parent, *, breakpoint=860):
        super().__init__(parent, corner_radius=0, fg_color='transparent')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.panes = tk.PanedWindow(self, orient='horizontal', borderwidth=0,
                                    sashwidth=8, opaqueresize=True, showhandle=False)
        self.panes.grid(row=0, column=0, sticky='nsew')
        self.left, self.right = (ctk.CTkFrame(self.panes, corner_radius=12) for _ in range(2))
        self.panes.add(self.left, minsize=260, width=340, stretch='never')
        self.panes.add(self.right, minsize=360, stretch='always')
        self.vertical, self.wide_sash, self.pending = None, 340, None
        self.breakpoint = breakpoint
        tk.Misc.bind(self, '<Configure>', self._resize, add='+')

    def _resize(self, event):
        if event.widget is self and self.pending is None:
            self.pending = self.after_idle(self._layout)

    def _layout(self):
        self.pending = None
        vertical = self.winfo_width()/self._get_widget_scaling()<self.breakpoint
        if vertical==self.vertical:
            return
        if self.vertical is False:
            self.wide_sash = self.panes.sash_coord(0)[0]
        self.vertical = vertical
        self.panes.configure(orient='vertical' if vertical else 'horizontal')
        self.panes.paneconfigure(self.left, minsize=110 if vertical else 260,
                                 width=340, height=180, stretch='never')
        self.panes.paneconfigure(self.right, minsize=170 if vertical else 360, stretch='always')
        self.panes.update_idletasks()
        self.panes.sash_place(0, 0 if vertical else self.wide_sash, 180 if vertical else 0)

    def apply_theme(self, colors):
        self.configure(fg_color=colors['bg'])
        self.panes.configure(background=colors['bg'])
        for pane in (self.left, self.right):
            pane.configure(fg_color=colors['surface'], bg_color=colors['bg'])

    def destroy(self):
        if self.pending is not None:
            self.after_cancel(self.pending)
        super().destroy()


class VirtualCard(ctk.CTkFrame):
    def __init__(self, parent, owner):
        super().__init__(parent, corner_radius=9, border_width=0)
        self.owner, self.index = owner, None
        self.grid_columnconfigure(0, weight=1)
        self.title = ctk.CTkLabel(self, anchor='w', height=24, font=font('section', True))
        self.title.grid(row=0, column=0, sticky='ew', padx=(14, 8), pady=(9, 0))
        self.robot = ctk.CTkLabel(self, height=22, corner_radius=5, width=40, font=font('meta', True))
        self.robot.grid(row=0, column=1, padx=(0, 12), pady=(9, 0))
        meta_size = 12 if getattr(owner.view, 'embedded', False) else SIZES['meta']
        self.robot.configure(font=('Segoe UI', meta_size, 'bold'))
        self.source = ctk.CTkLabel(self, anchor='w', height=20, font=('Segoe UI', meta_size))
        self.source.grid(row=1, column=0, columnspan=2, sticky='ew', padx=14, pady=(0, 9))
        if getattr(owner.view, 'compact_rows', False):
            self.title.configure(height=18, font=('Segoe UI', 12, 'bold'))
            self.title.grid_configure(padx=(8, 4), pady=(3, 0))
            self.robot.configure(height=18, font=('Segoe UI', 10, 'bold'))
            self.robot.grid_configure(padx=(0, 6), pady=(3, 0))
            self.source.configure(height=16, font=('Segoe UI', 11))
            self.source.grid_configure(padx=8, pady=(0, 3))
        self._font = ctk.CTkFont(family='Segoe UI', size=meta_size)
        self._content = None
        for widget in (self, self.title, self.robot, self.source):
            widget.configure(cursor='hand2')
            widget.bind('<ButtonRelease-1>', self._click, add='+')
            widget.bind('<Button-3>', self._context, add='+')
            widget.bind('<Double-Button-1>', lambda event: self.owner.view.activate(self.index), add='+')
            widget.bind('<MouseWheel>', owner._wheel, add='+')

    def _click(self, event):
        self.owner.canvas.focus_set()
        self.owner.view.select(self.index)
        return 'break'

    def _context(self, event):
        self.owner.canvas.focus_set()
        self.owner.view.context(self.index, event)
        return 'break'

    def show(self, index, width):
        self.index = index
        if self._content != (index, width):
            self._content = index, width
            content = self.owner.view.card_content(index)
            self.title.configure(text=content['title'])
            self.robot.configure(text=content['badge'])
            source = content['subtitle']
            while len(source)>2 and self._font.measure(source)>max(50, width-32):
                source = source[:-2].rstrip('…') + '…'
            self.source.configure(text=source)
        self.apply_theme()

    def apply_theme(self):
        colors = self.owner.view.app._palette
        active = self.index == self.owner.view.selected_index
        self.configure(fg_color=colors['info_bg'] if active else colors['bg'],
                       bg_color=colors['surface'], border_width=1 if active else 0, border_color=colors['accent'])
        self.title.configure(text_color=colors['text'])
        self.source.configure(text_color=colors['muted'])
        self.robot.configure(fg_color=colors['surface'], text_color=colors['accent'])


class VirtualCardList(ctk.CTkFrame):
    """Bounded pool of CTk cards: only the viewport and three extra rows exist."""
    def __init__(self, parent, view):
        super().__init__(parent, corner_radius=0, fg_color='transparent')
        self.view, self.cards = view, []
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, takefocus=True)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        self.scrollbar = ctk.CTkScrollbar(self, command=self.canvas.yview, width=12)
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        self.canvas.configure(yscrollcommand=self._scrolled)
        self.canvas.bind('<Configure>', self._resize)
        self.canvas.bind('<MouseWheel>', self._wheel)
        for key in ('Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
            self.canvas.bind('<'+key+'>', self._key)
        self.canvas.bind('<Control-c>', lambda event: view.copy('row'))
        self.canvas.bind('<Return>', lambda event: view.activate(view.selected_index))
        self._drawing = False
        self._draw_key = None

    @property
    def row_height(self):
        return round((48 if getattr(self.view, 'compact_rows', False) else 76)*self._get_widget_scaling())

    def _resize(self, event):
        self.canvas.configure(scrollregion=(0, 0, event.width,
            max(event.height, self.row_height*len(self.view.errors))))
        self.draw()

    def _scrolled(self, first, last):
        self.scrollbar.set(first, last)
        self.draw()

    def _wheel(self, event):
        self.canvas.yview_scroll(-int(event.delta/120)*2, 'units')
        return 'break'

    def _key(self, event):
        count = len(self.view.errors)
        if not count:
            return 'break'
        current = self.view.selected_index or 0
        page = max(1, self.canvas.winfo_height()//self.row_height-1)
        target = {'Up': current-1, 'Down': current+1, 'Home': 0,
                  'End': count-1, 'Prior': current-page, 'Next': current+page}[event.keysym]
        target = max(0, min(count-1, target))
        self.see(target)
        self.view.select(target)
        return 'break'

    def see(self, index):
        if index is None:
            return
        top, height = self.canvas.canvasy(0), self.canvas.winfo_height()
        start, end = index*self.row_height, (index+1)*self.row_height
        if start<top or end>top+height:
            destination = start if start<top else end-height
            self.canvas.yview_moveto(destination/max(1, self.row_height*len(self.view.errors)))
        self.draw()

    def draw(self):
        if self._drawing or not self.winfo_exists():
            return
        self._drawing = True
        try:
            width = max(100, self.canvas.winfo_width()-8)
            start = max(0, int(self.canvas.canvasy(0)//self.row_height)-1)
            count = min(len(self.view.errors)-start, math.ceil(self.canvas.winfo_height()/self.row_height)+3)
            key = start, count, width, self.view.selected_index, tuple(self.view.app._palette.items())
            if key == self._draw_key:
                return
            self._draw_key = key
            for offset in range(max(0, count)):
                if offset==len(self.cards):
                    card = VirtualCard(self.canvas, self)
                    item = self.canvas.create_window(4, 0, anchor='nw', window=card)
                    self.cards.append((card, item))
                card, item = self.cards[offset]
                index = start+offset
                card.show(index, width)
                self.canvas.itemconfigure(item, width=width, height=self.row_height-6, state='normal')
                self.canvas.coords(item, 4, index*self.row_height+2)
            for card, item in self.cards[max(0, count):]:
                self.canvas.itemconfigure(item, state='hidden')
        finally:
            self._drawing = False

    def apply_theme(self):
        colors = self.view.app._palette
        self.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        self.canvas.configure(background=colors['surface'])
        self.scrollbar.configure(fg_color=colors['surface'], button_color=colors['border'], button_hover_color=colors['muted'])
        for card, _ in self.cards:
            card.apply_theme()


class StateCard(ctk.CTkFrame):
    def __init__(self, parent, title):
        super().__init__(parent, corner_radius=10)
        self.role = 'neutral'
        self.grid_columnconfigure(0, weight=1)
        self.caption = selectable(self, text=title, font=font('meta', True))
        self.caption.grid(row=0, column=0, sticky='ew', padx=14, pady=(12, 8))
        self.value = selectable(self, text='', font=font(bold=True), wraplength=270)
        self.value.grid(row=1, column=0, sticky='ew', padx=14)
        self.time = selectable(self, text='', font=('Segoe UI', 11))
        self.time.grid(row=2, column=0, sticky='ew', padx=14, pady=(8, 12))

    def show(self, role, value, when=''):
        self.role = role
        self.value.configure(text=value)
        self.time.configure(text=when or '—')
        for label in (self.value, self.time):
            label.sync_selection_text()

    def apply_theme(self, colors):
        background, foreground = {'positive': ('positive_bg', 'positive_text'), 'info': ('info_bg', 'info_text'),
            'waiting': ('error_bg', 'error_text')}.get(self.role, ('raised', 'muted'))
        self.configure(fg_color=colors[background], bg_color=colors['surface'])
        for label in (self.caption, self.value, self.time):
            label.configure(text_color=colors[foreground], bg_color=colors[background])
            label.sync_selection_text()

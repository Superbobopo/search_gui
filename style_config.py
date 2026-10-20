from tkinter import ttk


def configure_style(root, palette=None):
    """Use the same colors for ttk and the native/custom application frame."""
    from ui_theme import LIGHT
    colors = palette or LIGHT
    from ui_components import font
    style = ttk.Style(root)
    if style.theme_use() != 'clam':
        style.theme_use('clam')
    style.configure('.', font=font(native=True), background=colors['bg'],
                    foreground=colors['text'], troughcolor=colors['raised'],
                    bordercolor=colors['border'], lightcolor=colors['border'],
                    darkcolor=colors['border'])
    for name, padding in [('TButton', (10, 6)), ('Nav.TButton', (8, 4)),
                          ('Arrow.TButton', (4, 2))]:
        style.configure(name, font=font(native=True), padding=padding,
                        borderwidth=1, relief='flat', focusthickness=2,
                        focuscolor=colors['accent'], background=colors['raised'],
                        foreground=colors['text'])
        style.map(name, background=[('disabled', colors['bg']), ('active', colors['border'])],
                  foreground=[('disabled', colors['muted']), ('active', colors['text'])],
                  bordercolor=[('focus', colors['accent'])])
    style.configure('TEntry', fieldbackground=colors['surface'],
                    foreground=colors['text'], insertcolor=colors['text'], padding=5)
    style.configure('TCheckbutton', background=colors['bg'])
    style.configure('TLabel', background=colors['bg'], foreground=colors['text'])
    style.configure('TLabelframe', background=colors['surface'])
    style.configure('TLabelframe.Label', background=colors['surface'], foreground=colors['text'])
    from tkinter import font as tkfont
    rowheight = max(34, tkfont.Font(root=root, font=('Segoe UI', 10)).metrics('linespace') + 14)
    style.configure('Tracking.Treeview', font=font(native=True), rowheight=rowheight,
                    background=colors['surface'], fieldbackground=colors['surface'],
                    foreground=colors['text'], borderwidth=0, relief='flat')
    style.configure('Tracking.Treeview.Heading', font=font('meta', True, native=True),
                    background=colors['raised'], foreground=colors['muted'], padding=(10, 9),
                    relief='flat', borderwidth=0)
    style.map('Tracking.Treeview', background=[('selected', colors['accent'])],
              foreground=[('selected', colors['accent_text'])])
    return style

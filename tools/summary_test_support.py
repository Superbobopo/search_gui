"""Selectors for the single search presentation in isolated GUI checks."""
import customtkinter as ctk


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def rendered_summary_text(app):
    return '\n'.join(str(w.cget('text')) for w in descendants(app.summary_frame)
                     if isinstance(w, ctk.CTkLabel))


def click_summary(app, key, stamp=1000, double=True):
    key = key.removeprefix('clickable_')
    if key=='refprod':
        field = next(w for w in descendants(app.summary_frame)
                     if hasattr(w, 'selection_text') and w.cget('text')==app.selection.reference)
    else:
        field = app.summary_targets[key]
    widget = field.selection_text
    canvas = app.summary_frame._parent_canvas
    app.update_idletasks()
    top = canvas.canvasy(widget.winfo_rooty()-canvas.winfo_rooty())
    region = canvas.bbox('all')
    if region:
        canvas.yview_moveto(max(0, top-30)/max(1, region[3]))
    app.update_idletasks()
    widget.event_generate('<Motion>', x=5, y=5)
    for i in range(2 if double else 1):
        widget.event_generate('<ButtonPress-1>', x=5, y=5, time=stamp+i*100)
        widget.event_generate('<ButtonRelease-1>', x=5, y=5, time=stamp+i*100+20)

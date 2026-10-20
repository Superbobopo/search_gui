"""Modeless, independently cancellable GET failure report."""
from datetime import datetime
import os
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk

from background_tasks import BackgroundTasks
from popup_style import PopupTreeview, configure_popup_style
from trace_sources import compact_trace_label
from ui_components import WindowShell, ActionButton, selectable, font, style_labels


def error_key(row):
    return str(row['path']),getattr(row['path'],'version_id',None),row['number'],bool(row['get'])


class GetErrorsView(WindowShell):
    columns = ('date','hour','robot','location','ref','modbus','error','trace')

    def __init__(self,parent,app):
        super().__init__(parent,app,'Erreurs GET','Analyse des erreurs GET…')
        self.tasks = BackgroundTasks()
        self.data, self.rows, self.values = None, {}, {}
        self._closed, self._loading = False, False
        self._marker, self._request_marker = None, None
        self._failed_marker = None
        self._poll_timer = self._insert_timer = None
        self._sort = ('time',True)
        self._selected_identity = None
        self.form = ctk.CTkFrame(self.body,fg_color='transparent')
        self.form.grid(row=0,column=0,sticky='ew',padx=12,pady=10)
        self.body.grid_rowconfigure(0,weight=0)
        self.body.grid_rowconfigure(1,weight=1)
        self.form.grid_columnconfigure((0,1),weight=1)
        self.filters = []
        for index,caption in enumerate(('Robot','Emplacement')):
            field = ctk.CTkFrame(self.form,fg_color='transparent')
            field.grid(row=0,column=index,sticky='ew',padx=(0,12))
            field.grid_columnconfigure(0,weight=1)
            ctk.CTkLabel(field,text=caption,font=font('meta'),anchor='w').grid(row=0,column=0,sticky='w')
            choice = ctk.CTkOptionMenu(field,values=['Tous'],command=lambda value:self._display())
            choice.grid(row=1,column=0,sticky='ew')
            self.filters.append(choice)
        self.robot_filter,self.location_filter = self.filters
        self.refresh_button = ActionButton(self.form,text='Actualiser',command=self.refresh)
        self.refresh_button.grid(row=0,column=2,sticky='s',pady=(16,0))
        table = self.table = ctk.CTkFrame(self.body,fg_color='transparent')
        table.grid(row=1,column=0,sticky='nsew',padx=12)
        table.grid_rowconfigure(0,weight=1); table.grid_columnconfigure(0,weight=1)
        self.tree = PopupTreeview(table,columns=self.columns,show='headings',selectmode='browse',style='GetErrors.Treeview')
        for key,title,width in zip(self.columns,('DATE','HEURE','ROBOT','EMPLACEMENT','RÉFÉRENCE','ID','ERREUR','TRACE'),
                                    (90,90,60,105,140,65,180,200)):
            self.tree.heading(key,text=title,command=lambda key=key:self.sort(key))
            self.tree.column(key,width=width,minwidth=55 if key in ('robot','modbus') else 85,
                             stretch=key in ('ref','trace'),anchor='center' if key in ('date','hour','robot','modbus') else 'w')
        self.tree.grid(row=0,column=0,sticky='nsew')
        for orientation,command,row,column,sticky in (('vertical',self.tree.yview,0,1,'ns'),('horizontal',self.tree.xview,1,0,'ew')):
            bar = ttk.Scrollbar(table,orient=orientation,command=command,
                                style=f'Popup.{orientation.title()}.TScrollbar')
            bar.grid(row=row,column=column,sticky=sticky)
            self.tree.configure(**{'yscrollcommand' if orientation=='vertical' else 'xscrollcommand':bar.set})
        self.detail = ctk.CTkFrame(self.body,fg_color='transparent')
        self.detail.grid(row=2,column=0,sticky='ew',padx=12,pady=10)
        self.detail.grid_columnconfigure(0,weight=1)
        self.detail_title = selectable(self.detail,text='Sélectionnez une erreur.',font=font('section',True),wraplength=900)
        self.identity = selectable(self.detail,text='',font=font('meta'),wraplength=900)
        self.source = selectable(self.detail,text='',font=font('meta'),wraplength=900)
        self.command = selectable(self.detail,text='',font=('Consolas',11),wraplength=900)
        self.responses = selectable(self.detail,text='',font=('Consolas',11),wraplength=900)
        for index,widget in enumerate((self.detail_title,self.identity,self.source,self.command,self.responses)):
            widget.grid(row=index,column=0,sticky='ew',pady=2)
        self.open_button = ActionButton(self.detail,text='Ouvrir la trace',role='quiet',command=self.open_error)
        self.open_button.grid(row=5,column=0,sticky='w',pady=(4,0))
        self.tree.bind('<<TreeviewSelect>>',self._select)
        self.tree.bind('<Double-Button-1>',self.open_error)
        self.tree.bind('<Return>',self.open_error)
        self.tree.bind('<Control-c>',self.copy_row)
        self.tree.bind('<Button-3>',self._menu)
        self.command.selection_text.bind('<Double-Button-1>',self.open_command,add='+')
        self.responses.selection_text.bind('<Double-Button-1>',self.open_error,add='+')
        self.menu = tk.Menu(self,tearoff=False)
        self.menu.add_command(label='Copier la cellule',command=self.copy_cell)
        self.menu.add_command(label='Copier la ligne',command=self.copy_row)
        self._cell = None
        tk.Misc.bind(self.detail,'<Configure>',self._resize_detail,add='+')
        tk.Misc.bind(self,'<Destroy>',self._destroyed,add='+')
        self.apply_theme()
        self.refresh()
        self._poll_timer = self.after(30,self._poll)

    def _ready(self):
        return self.app._sources_verified and self.app._indexed_complete

    def refresh(self):
        if self._closed or not self._ready():
            return
        self._loading = True
        self._failed_marker = None
        self._request_marker = self.app._index_content_revision
        marker = self._request_marker
        self.header.subtitle.configure(text='Analyse des erreurs GET…')
        self.open_button.configure(state='disabled')
        self.refresh_button.configure(state='disabled')

        def work(cancel):
            return self.app.store.get_errors(cancel)

        def done(data):
            if not self._ready() or marker!=self.app._index_content_revision:
                self._request_marker = None
                return
            self._loading = False
            self._marker = marker
            self._request_marker = None
            self.data = data
            self.refresh_button.configure(state='normal')
            for widget,values in ((self.robot_filter,['Tous','R1','R2']+(['Non identifié'] if any(not r['robot'] for r in data['errors']) else [])),
                                  (self.location_filter,['Tous']+sorted({r['location'] for r in data['errors'] if r['location']})
                                   +(['Inconnu'] if any(not r['location'] for r in data['errors']) else []))):
                current = widget.get()
                if current not in values:
                    values.append(current)
                widget.configure(values=values)
            self._display()

        def failed(error):
            self._loading = False
            self._request_marker = None
            self.header.subtitle.configure(text=f'Analyse interrompue : {error}. Actualisez pour réessayer.')
            self.refresh_button.configure(state='normal')
            # Previous rows remain visible, but cannot open an unvalidated result.
            self._marker = None
            self._failed_marker = marker

        self.tasks.submit('get_errors',work,done,failed)

    def observe_index(self):
        if not self._ready():
            self.tasks.cancel('get_errors')
            self._request_marker = None
            self._marker = None
            self.header.subtitle.configure(text='Disponible après vérification complète des sources.')
            self.refresh_button.configure(state='disabled')
            self.open_button.configure(state='disabled')
        elif (self._marker!=self.app._index_content_revision and self._request_marker is None
              and self._failed_marker!=self.app._index_content_revision):
            self.refresh()

    def _poll(self):
        self._poll_timer = None
        if self._closed:
            return
        self.observe_index()
        self.tasks.drain()
        self._poll_timer = self.after(30,self._poll)

    def _display(self):
        if self.data is None:
            return
        if self._insert_timer is not None:
            self.after_cancel(self._insert_timer); self._insert_timer = None
        scroll = self.tree.yview()[0]
        selected = self._selected_identity
        robot,location = self.robot_filter.get(),self.location_filter.get()
        rows = [r for r in self.data['errors'] if (robot=='Tous' or (r['robot'] or 'Non identifié')==robot)
                and (location=='Tous' or (r['location'] or 'Inconnu')==location)]
        key,reverse = self._sort
        def order(row):
            if key=='time': return row['time']
            if key=='error': return row['message'].casefold()
            if key=='trace': return str(row['path']).casefold()
            return str(row.get(key) or '').casefold()
        rows.sort(key=order,reverse=reverse)
        count = self._visible_count = len(rows)
        traces = len({(str(r['path']),getattr(r['path'],'version_id',None)) for r in rows})
        suffix = 's' if count!=1 else ''
        trace_suffix = 's' if traces!=1 else ''
        self.header.subtitle.configure(text=f'{count} erreur{suffix} affichée{suffix} / {len(self.data["errors"])} · '
                                           f'{traces} trace{trace_suffix} concernée{trace_suffix}')
        items = self.tree.get_children()
        if items: self.tree.delete(*items)
        self.rows,self.values = {},{}
        self._show_detail(None)
        wanted = next((i for i,r in enumerate(rows) if error_key(r)==selected),0)
        iterator = iter(enumerate(rows))
        def append():
            self._insert_timer = None
            for _ in range(200):
                item = next(iterator,None)
                if item is None:
                    self.tree.yview_moveto(scroll)
                    return
                index,row = item; iid = str(index)
                when = datetime.fromtimestamp(row['time'])
                values = (when.strftime('%d/%m/%Y'),when.strftime('%H:%M:%S.%f').rstrip('0').rstrip('.'),
                          row['robot'] or '—',row['location'] or '—',row['ref'] or '—',row['modbus'] or '—',
                          row['message'],compact_trace_label(row['path'],self.app._selected_directories))
                self.rows[iid],self.values[iid] = row,values
                self.tree.insert('','end',iid=iid,values=values,tags=('odd' if index%2 else 'even',))
                if index==wanted:
                    self.tree.selection_set(iid); self.tree.focus(iid); self._show_detail(row)
            self._insert_timer = self.after(1,append)
        append()

    def sort(self,key):
        key = 'time' if key in ('date','hour') else key
        self._sort = (key,not self._sort[1] if self._sort[0]==key else False)
        self._display()

    def _selected(self,event=None):
        iid = self.tree.identify_row(event.y) if event is not None and event.widget is self.tree and getattr(event,'type',None)==tk.EventType.ButtonPress else self.tree.focus()
        return self.rows.get(iid)

    def _select(self,event=None):
        self._show_detail(self._selected())

    def _show_detail(self,row):
        if row:
            self._selected_identity = error_key(row)
            when = datetime.fromtimestamp(row['time']).strftime('%d/%m/%Y %H:%M:%S')
            title = 'Échec GET — '+row['message']
            identity = f"{row['robot'] or 'Robot non identifié'} · Face {row['face'] or '—'} · {row['location'] or 'Emplacement inconnu'} · {when}"
            identity += f"\nRéférence : {row['ref'] or '—'} · ID : {row['modbus'] or '—'} · SN : {row['sn'] or '—'} · CodeLot : {row['lot'] or '—'}"
            source = compact_trace_label(row['path'],self.app._selected_directories)+f" · Ligne {row['error']['number']+1}"
            command = row['get']['text'] if row['get'] else 'Commande GET non rattachée.'
            responses = '\n'.join(r['text'] for r in row['responses'])
        else:
            title = 'Aucune erreur GET pour ces filtres.' if getattr(self,'_visible_count',None)==0 else 'Sélectionnez une erreur.'
            identity,source,command,responses = '','','',''
        for widget,value in zip((self.detail_title,self.identity,self.source,self.command,self.responses),
                                (title,identity,source,command,responses)):
            if widget.cget('text')!=value:
                widget.configure(text=value); widget.sync_selection_text()
        self.open_button.configure(state='normal' if row and self._can_open() else 'disabled')

    def _can_open(self):
        return not self._loading and self._ready() and self._marker==self.app._index_content_revision

    def _open(self,row,event):
        if self._can_open() and row and event:
            self.app._open_indexed_file(row['path'],event['number'],event['text'])
        return 'break'

    def open_error(self,event=None):
        row = self._selected(event)
        return self._open(row,row['error'] if row else None)

    def open_command(self,event=None):
        row = self._selected()
        return self._open(row,row['get'] if row else None)

    def copy_row(self,event=None):
        value = self.values.get(self.tree.focus())
        if value:
            self.clipboard_clear(); self.clipboard_append(' | '.join(value))
        return 'break'

    def copy_cell(self):
        if self._cell:
            iid,key = self._cell
            self.clipboard_clear(); self.clipboard_append(self.values[iid][self.columns.index(key)])

    def _menu(self,event):
        iid,column = self.tree.identify_row(event.y),self.tree.identify_column(event.x)
        if iid not in self.rows or not column or column=='#0': return
        self._cell = iid,self.tree.column(column,'id')
        self.tree.focus(iid); self.tree.selection_set(iid)
        try: self.menu.tk_popup(event.x_root,event.y_root)
        finally: self.menu.grab_release()

    def _resize_detail(self,event):
        width = max(180,round(event.width/self._get_widget_scaling())-24)
        for widget in (self.detail_title,self.identity,self.source,self.command,self.responses):
            if widget.cget('wraplength')!=width: widget.configure(wraplength=width)

    def apply_theme(self):
        WindowShell.apply_theme(self)
        if not hasattr(self,'tree'): return
        colors = self.app._palette
        style = configure_popup_style(self,colors)
        style.map('GetErrors.Treeview',background=[('selected',colors['info_bg'])],foreground=[('selected',colors['warning_text'])])
        for tag,color in (('even',colors['surface']),('odd',colors['bg'])):
            self.tree.tag_configure(tag,background=color,foreground=colors['warning_text'])
        for pane in (self.form,self.table,self.detail): pane.configure(fg_color=colors['surface'],bg_color=colors['surface'])
        style_labels((self.detail_title,self.responses),colors,'warning_text')
        style_labels((self.identity,self.source,self.command),colors)
        for button in (self.refresh_button,self.open_button): button.apply_theme(colors)

    def _destroyed(self,event):
        if event.widget is self: self._dispose()

    def _dispose(self):
        if self._closed: return
        self._closed = True
        for timer in (self._poll_timer,self._insert_timer):
            if timer is not None: self.after_cancel(timer)
        self._poll_timer = self._insert_timer = None
        self.tasks.close()

    def destroy(self):
        self._dispose()
        super().destroy()

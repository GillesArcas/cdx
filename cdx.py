from __future__ import print_function

import sqlite3
import sys
import os
import tkinter as tk
try:
    import tkFileDialog as tkfiledialog
except:
    import tkinter.filedialog as tkfiledialog
import tempfile
import base64
import zlib
import time
import string
try:
    from win32com.shell import shell
except:
    pass  # on a deja tkfiledialog


# -- Path database


class PathManager(object):

    def __init__(self):
        self.db_name = os.path.splitext(sys.argv[0])[0] + '.sqlite3'
        self.conn = sqlite3.connect(self.db_name)
        self.conn.execute('CREATE TABLE IF NOT EXISTS paths (path TEXT, count INT, time TEXT)')
        self.conn.commit()

    def __del__(self):
        if self.conn:
            self.conn.close()

    def use_path(self, path):
        path = normalize_path(path)
        for path, count, _ in self.conn.execute('SELECT * FROM paths WHERE path=?', (path,)):
            self.conn.execute('UPDATE paths SET count=?,time=? WHERE path=?', (count + 1, cdtime(), path))
            self.conn.commit()
            break

    def add_path(self, path):
        if not path:
            return

        path = normalize_path(path)

        if list(self.conn.execute('SELECT * FROM paths WHERE path=?', (path,))):
            return

        try:
            self.conn.execute('INSERT INTO paths (path, count, time) VALUES (?,?,?)', (path, 0, cdtime()))
            self.conn.commit()
        except sqlite3.Error as error:
            tk.messagebox.showerror('Path manager', error)

    def del_path(self, path):
        try:
            self.conn.execute('DELETE FROM paths WHERE path=?', (path,))
            self.conn.commit()
        except sqlite3.Error as error:
            tk.messagebox.showerror('Path manager', error)

    def paths(self, sorton=0, sortdir=0):
        sorton_field = ('path', 'count', 'time')
        sortdir_kw = ('ASC', 'DESC')
        com = 'SELECT * FROM paths ORDER BY %s %s' % (sorton_field[sorton], sortdir_kw[sortdir])
        return list(self.conn.execute(com))


# -- Interface


ICON = r"""\
eJxjYGAEQgEBASDJwqDByMAgxsDAoAHEAkCsAMQgcRBoYMAO/v//j0OGfMAoCAZAu8E80mmofhAQgI
rCTYOKQ9yOwA3MDAz7f+PGIPn/1UBsDcTNED0AZe8evQ==
"""

class PathManagerApp(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        w = 500
        h = 200
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.title('Path Manager')

        with tempfile.NamedTemporaryFile(suffix='.ico', delete=False) as f:
            f.write(zlib.decompress(base64.b64decode(ICON)))
            f.close()
            self.master.iconbitmap(f.name)
            os.remove(f.name)

        self.grid(sticky='nesw')
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=0)
        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=1)

        self.menu = MenuBar(master=self)
        self.menu.grid(row=0, column=0)

        self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.scrollbar.grid(row=1, column=1, sticky='nes')

        self.path_listbox = tk.Listbox(master=self,
                                       selectmode=tk.SINGLE,
                                       yscrollcommand=self.scrollbar.set)
        self.path_listbox.bind('<Double-Button-1>', self.list_click_handler)
        self.path_listbox.grid(row=1, column=0, sticky='nsew')
        self.path_listbox.config(font=("Lucida Console", 9, ""))

        self.scrollbar.config(command=self.path_listbox.yview)

        self.path_manager = PathManager()
        self.show_paths()

    def show_paths(self, sorton=0, sortdir=0):
        self.path_listbox.delete(0, tk.END)
        for path, count, time in self.path_manager.paths(sorton, sortdir):
            self.path_listbox.insert(tk.END, format_listbox_line(path, count, time))

    def list_click_handler(self, event):
        selected_row = self.path_listbox.curselection()
        line = self.path_listbox.get(selected_row)
        path, _, _ = parse_listbox_line(line)
        self.path_manager.use_path(path)
        print(path)
        sys.exit(0)


class MenuBar(tk.Frame):

    def __init__(self, master=None):
        self.master = master
        tk.Frame.__init__(self, master)
        self.grid(sticky=tk.W)
        widget = tk.Button(self, text="Select", relief=tk.FLAT, command=self.on_click_select)
        widget.grid(row=0, column=0)
        widget = tk.Button(self, text="Remove", relief=tk.FLAT, command=self.on_click_remove)
        widget.grid(row=0, column=1)

        widget = tk.Label(self, text='Sort:')
        widget.grid(row=0, column=2)
        self.var_sorton = tk.IntVar()
        self.var_sorton.set(0)
        self.create_radio('name', self.var_sorton, val=0, col=3)
        self.create_radio('count', self.var_sorton, val=1, col=4)
        self.create_radio('date', self.var_sorton, val=2, col=5)
        self.var_sortdir = tk.IntVar()
        self.var_sortdir.set(0)
        self.create_radio('asc', self.var_sortdir, val=0, col=6)
        self.create_radio('desc', self.var_sortdir, val=1, col=7)

    def create_radio(self, text, var, val, col):
        button = tk.Radiobutton(self, text=text, variable=var, value=val, command=self.onclick_sort)
        button.grid(row=0, column=col)

    def on_click_select(self):
        selected_row = self.master.path_listbox.curselection()
        if selected_row:
            line = self.master.path_listbox.get(selected_row)
            initialdir, _, _ = parse_listbox_line(line)
        else:
            initialdir = 'd:\\'
        path = selectdir(initialdir)
        if path:
            self.master.path_manager.add_path(path)
            self.master.path_manager.use_path(path)
            self.master.show_paths()
            print(path)
            sys.exit(0)

    def on_click_remove(self):
        selected_row = self.master.path_listbox.curselection()
        if selected_row:
            line = self.master.path_listbox.get(selected_row)
            path, _, _ = parse_listbox_line(line)
            self.master.path_manager.del_path(path)
            self.master.show_paths()

    def onclick_sort(self):
        self.master.show_paths(self.var_sorton.get(), self.var_sortdir.get())


def format_listbox_line(path, count, timestamp):
    return '%-40s | %3d | %s' % (path, count, timestamp)


def parse_listbox_line(line):
    return (_.strip() for _ in line.split('|'))


# -- Helpers


def selectdir(initialdir='d:/'):
    try:
        return selectdir_win(initialdir)
    except:
        return selectdir_tk(initialdir)


def selectdir_win(initialdir):
    title = 'Select directory'
    pidl = shell.SHILCreateFromPath(initialdir, 0)[0]
    pidl, display_name, image_list = shell.SHBrowseForFolder(None, pidl, title, 0, None, None)

    if (pidl, display_name, image_list) == (None, None, None):
        return None
    else:
        path = shell.SHGetPathFromIDList(pidl)
        return path


def selectdir_tk(initialdir):
    """
    on windows, does not scroll to show initial directory.
    """
    title = 'Select directory'
    filename = tkfiledialog.askdirectory(initialdir=initialdir, title=title, mustexist=True)
    return filename


def normalize_path(path):
    path = os.path.normpath(path)
    if path[1] == ':':
        if path[0] in string.ascii_lowercase:
            path = path[0].upper() + path[1:]
    return path


def cdtime():
    return time.strftime("%Y-%m-%d %H:%M")


# -- Main


def insert_path(path):
    if path != '.' and path != '..':
        path_manager = PathManager()
        path_manager.add_path(path)
        path_manager.use_path(path)


if __name__ == '__main__':
    if sys.argv[1] == 'insert':
        insert_path(sys.argv[2])
    elif sys.argv[1] == 'select':
        PathManagerApp().mainloop()
    else:
        print('Appel cdx.py incorrect:', ' '.join(sys.argv))

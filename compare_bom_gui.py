import os
import re
import shutil
import threading
import tkinter as tk
import zipfile
from io import BytesIO
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from tkinterdnd2 import TkinterDnD, DND_FILES

YELLOW_FILL = PatternFill(fill_type='solid', start_color='FFFFFF00', end_color='FFFFFF00')
RED_COLOR = 'FFFF0000'
MAX_COL = 9  # Fill yellow only up to column I
DATA_SHEETS = ['LM', 'LP', 'LS', 'LW', 'LB']


# ── Comparison logic ─────────────────────────────────────────────────────────

def vals_equal(a, b):
    a_none = a is None or (isinstance(a, float) and pd.isna(a))
    b_none = b is None or (isinstance(b, float) and pd.isna(b))
    if a_none and b_none:
        return True
    if a_none or b_none:
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        fa, fb = float(a), float(b)
        diff = abs(fa - fb)
        mag = max(abs(fa), abs(fb), 1e-12)
        return diff / mag < 1e-6
    return str(a).strip() == str(b).strip()


def is_pozycja(val):
    if val is None:
        return False
    return bool(re.match(r'^[A-Z]{1,4}-\d+', str(val).strip()))


def is_lb_data_row(vals):
    """LB data row: any non-empty col A that is not the header line."""
    v = vals[0]
    if v is None:
        return False
    s = str(v).strip()
    return s != '' and s != 'Pozycja:'


def is_parent_ls(vals):
    if not is_pozycja(vals[0]):
        return False
    b_empty = len(vals) < 2 or vals[1] is None or str(vals[1]).strip() == ''
    c_empty = len(vals) < 3 or vals[2] is None or str(vals[2]).strip() == ''
    return b_empty and c_empty


def read_data_rows(wb_data_only, sheet_name, data_start):
    ws = wb_data_only[sheet_name]
    rows = []
    for excel_row in range(data_start, ws.max_row + 1):
        vals = [ws.cell(row=excel_row, column=c).value
                for c in range(1, ws.max_column + 1)]
        rows.append((excel_row, vals))
    return rows


def build_old_map(old_rows, is_ls=False, is_lb=False):
    """
    Returns dict: key -> list of vals (ordered by appearance).
    For LS: key = (pozycja, 'parent'|'child').
    For LB: key = pozycja string (any col-A value), duplicates stored in order.
    For others: key = pozycja string.
    """
    result = {}
    for _excel_row, vals in old_rows:
        if is_lb:
            if not is_lb_data_row(vals):
                continue
            key = str(vals[0]).strip()
        else:
            if not is_pozycja(vals[0]):
                continue
            key = str(vals[0]).strip()
            if is_ls:
                key = (key, 'parent' if is_parent_ls(vals) else 'child')
        if key not in result:
            result[key] = []
        result[key].append(vals)
    return result


def build_ls_child_to_parent(rows):
    child_to_parent = {}
    current_parent = None
    for _r, vals in rows:
        if not is_pozycja(vals[0]):
            continue
        if is_parent_ls(vals):
            current_parent = str(vals[0]).strip()
        elif current_parent:
            child_to_parent[str(vals[0]).strip()] = current_parent
    return child_to_parent


def apply_yellow_row(ws, excel_row, max_col):
    for c in range(1, max_col + 1):
        ws.cell(row=excel_row, column=c).fill = YELLOW_FILL


def apply_red_text(ws, excel_row, col_1based):
    if col_1based > MAX_COL:
        return
    cell = ws.cell(row=excel_row, column=col_1based)
    f = cell.font
    cell.font = Font(
        name=f.name, size=f.size, bold=True, italic=f.italic,
        underline=f.underline, strike=f.strike, vertAlign=f.vertAlign,
        color=RED_COLOR
    )


def compare_sheet(ws_out, old_rows, new_rows, is_ls=False, is_lb=False):
    old_map = build_old_map(old_rows, is_ls=is_ls, is_lb=is_lb)
    # occurrence counter: for duplicate keys, match nth new to nth old
    occurrence = {}
    changed_positions = set()

    for excel_row, new_vals in new_rows:
        if is_lb:
            if not is_lb_data_row(new_vals):
                continue
            pozycja = str(new_vals[0]).strip()
            key = pozycja
        else:
            if not is_pozycja(new_vals[0]):
                continue
            pozycja = str(new_vals[0]).strip()
            parent = is_ls and is_parent_ls(new_vals)
            key = (pozycja, 'parent' if parent else 'child') if is_ls else pozycja

        idx = occurrence.get(key, 0)
        occurrence[key] = idx + 1

        if key not in old_map or idx >= len(old_map[key]):
            # New item not present in old version
            apply_yellow_row(ws_out, excel_row, MAX_COL)
            changed_positions.add(pozycja)
            continue

        old_vals = old_map[key][idx]
        n_cols = max(len(new_vals), len(old_vals))
        changed_cols = []
        for i in range(n_cols):
            nv = new_vals[i] if i < len(new_vals) else None
            ov = old_vals[i] if i < len(old_vals) else None
            if not vals_equal(nv, ov):
                changed_cols.append(i + 1)

        if changed_cols:
            apply_yellow_row(ws_out, excel_row, MAX_COL)
            for col in changed_cols:
                apply_red_text(ws_out, excel_row, col)
            changed_positions.add(pozycja)

    return changed_positions


def mark_ls_parents(ws_out, new_rows, changed_children, child_to_parent):
    parents_to_mark = {child_to_parent[c] for c in changed_children if c in child_to_parent}
    if not parents_to_mark:
        return
    for excel_row, vals in new_rows:
        if is_pozycja(vals[0]) and is_parent_ls(vals):
            if str(vals[0]).strip() in parents_to_mark:
                apply_yellow_row(ws_out, excel_row, MAX_COL)


def clear_print_area_xml(xlsx_path):
    """
    Remove all Print_Area defined names directly from workbook.xml inside the xlsx zip.
    openpyxl re-generates these on save, so we must strip them post-save.
    """
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    tmp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
         zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                tree = ET.fromstring(data)
                defined_names = tree.find(f'{{{NS}}}definedNames')
                if defined_names is not None:
                    to_remove = [
                        dn for dn in defined_names
                        if 'Print_Area' in dn.get('name', '')
                    ]
                    for dn in to_remove:
                        defined_names.remove(dn)
                    # Remove empty definedNames element too
                    if len(defined_names) == 0:
                        tree.remove(defined_names)
                data = ET.tostring(tree, encoding='UTF-8', xml_declaration=True)
            zout.writestr(item, data)

    os.replace(tmp_path, xlsx_path)


def run_comparison(old_file, new_file, data_start, log_cb):
    out_dir = os.path.dirname(os.path.abspath(new_file))
    base_name = os.path.splitext(os.path.basename(new_file))[0]
    out_file = os.path.join(out_dir, base_name + '_POROWNANIE.xlsx')

    log_cb(f'Kopiowanie nowego pliku → {os.path.basename(out_file)}')
    shutil.copy2(new_file, out_file)

    wb_old_d = load_workbook(old_file, data_only=True)
    wb_new_d = load_workbook(new_file, data_only=True)
    wb_out = load_workbook(out_file)

    total_changed = 0
    for sheet in DATA_SHEETS:
        if sheet not in wb_old_d.sheetnames or sheet not in wb_out.sheetnames:
            log_cb(f'  {sheet}: pominięty (brak w jednym z plików)')
            continue

        old_rows = read_data_rows(wb_old_d, sheet, data_start)
        new_rows = read_data_rows(wb_new_d, sheet, data_start)
        ws_out = wb_out[sheet]

        is_ls = sheet == 'LS'
        is_lb = sheet == 'LB'
        changed = compare_sheet(ws_out, old_rows, new_rows, is_ls=is_ls, is_lb=is_lb)

        if is_ls:
            child_to_parent = build_ls_child_to_parent(new_rows)
            changed_children = set()
            for _r, vals in new_rows:
                if is_pozycja(vals[0]) and not is_parent_ls(vals):
                    if str(vals[0]).strip() in changed:
                        changed_children.add(str(vals[0]).strip())
            mark_ls_parents(ws_out, new_rows, changed_children, child_to_parent)

        log_cb(f'  {sheet}: {len(changed)} zmienione pozycje')
        total_changed += len(changed)

    log_cb('Zapisywanie...')
    wb_out.save(out_file)

    log_cb('Czyszczenie obszarów wydruku...')
    clear_print_area_xml(out_file)

    log_cb(f'\nGotowe! Łącznie zmian: {total_changed}')
    log_cb(f'Plik: {out_file}')
    return out_file


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title('BOM Comparer')
        self.resizable(False, False)
        self.configure(padx=20, pady=16)

        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.start_row = tk.StringVar(value='11')

        self._build_ui()

    def _build_ui(self):
        HINT = 'Przeciągnij plik lub kliknij aby wybrać ścieżkę'

        # ── File drop zones ──
        for label_text, var, attr in [
            ('Stara lista', self.old_path, 'old_drop'),
            ('Nowa lista',  self.new_path, 'new_drop'),
        ]:
            frame = tk.LabelFrame(self, text=label_text, padx=8, pady=8)
            frame.pack(fill='x', pady=4)

            drop = tk.Label(
                frame,
                text=HINT,
                relief='groove', bg='#f0f0f0',
                width=56, height=3,
                wraplength=420, justify='center',
                anchor='center'
            )
            drop.pack(side='left', fill='x', expand=True, padx=(0, 6))

            clear_btn = tk.Button(
                frame, text='✕', width=3,
                fg='#c62828', relief='flat', cursor='hand2',
                command=lambda v=var, d=drop, h=HINT: self._clear_path(v, d, h)
            )
            clear_btn.pack(side='left')

            # Bind drop and click
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind('<<Drop>>', lambda e, v=var, d=drop: self._on_drop(e, v, d))
            drop.bind('<Button-1>', lambda e, v=var, d=drop: self._browse(v, d))

            setattr(self, attr, drop)

        # ── Start row ──
        row_frame = tk.Frame(self)
        row_frame.pack(fill='x', pady=6)
        tk.Label(row_frame, text='Pierwszy wiersz danych (nagłówek +1):').pack(side='left')
        tk.Spinbox(row_frame, from_=2, to=999, textvariable=self.start_row,
                   width=6).pack(side='left', padx=8)

        # ── Run button ──
        self.run_btn = tk.Button(self, text='▶  Porównaj', command=self._run,
                                 bg='#2e7d32', fg='white', font=('Arial', 11, 'bold'),
                                 padx=16, pady=6, relief='flat', cursor='hand2')
        self.run_btn.pack(pady=10)

        # ── Log output ──
        log_frame = tk.LabelFrame(self, text='Postęp', padx=8, pady=8)
        log_frame.pack(fill='both', expand=True, pady=4)
        self.log = tk.Text(log_frame, height=10, state='disabled',
                           font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
                           relief='flat', wrap='word')
        self.log.pack(fill='both', expand=True)
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log['yscrollcommand'] = sb.set

    def _clear_path(self, var, label, hint):
        var.set('')
        label.config(text=hint, bg='#f0f0f0')

    def _on_drop(self, event, var, label):
        path = event.data.strip().strip('{}')  # Windows wraps paths in {}
        if path.lower().endswith('.xlsx'):
            var.set(path)
            label.config(text=os.path.basename(path) + '\n' + path, bg='#e8f5e9')
        else:
            messagebox.showerror('Błąd', 'Proszę przeciągnąć plik .xlsx')

    def _browse(self, var, label):
        path = filedialog.askopenfilename(
            filetypes=[('Excel files', '*.xlsx')],
            title='Wybierz plik Excel'
        )
        if path:
            var.set(path)
            label.config(text=os.path.basename(path) + '\n' + path, bg='#e8f5e9')

    def _log(self, msg):
        self.log.config(state='normal')
        self.log.insert('end', msg + '\n')
        self.log.see('end')
        self.log.config(state='disabled')

    def _run(self):
        old = self.old_path.get()
        new = self.new_path.get()
        if not old or not os.path.isfile(old):
            messagebox.showerror('Błąd', 'Wybierz starą listę materiałową.')
            return
        if not new or not os.path.isfile(new):
            messagebox.showerror('Błąd', 'Wybierz nową listę materiałową.')
            return
        try:
            start = int(self.start_row.get())
            if start < 2:
                raise ValueError
        except ValueError:
            messagebox.showerror('Błąd', 'Wiersz startowy musi być liczbą ≥ 2.')
            return

        self.run_btn.config(state='disabled', text='Przetwarzanie...')
        self.log.config(state='normal')
        self.log.delete('1.0', 'end')
        self.log.config(state='disabled')

        def task():
            try:
                run_comparison(old, new, start, self._log)
            except Exception as ex:
                self._log(f'\n❌ Błąd: {ex}')
            finally:
                self.run_btn.config(state='normal', text='▶  Porównaj')

        threading.Thread(target=task, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()

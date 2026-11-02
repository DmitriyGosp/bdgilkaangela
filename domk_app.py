# -*- coding: utf-8 -*-
"""Автозаповнення ДОМК-С-1 з Excel. Працює локально на Windows."""
import os, sys, calendar, traceback
from datetime import datetime, date, time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
from copy import copy

PASSENGER = {10, 13, 20, 21}
FREIGHT = {14, 22, 23, 24}
BUS = {9, 29}
# Межі звітних періодів: день 09:00:00–19:59:59; ніч 20:00:00–08:59:59.
DAY_START = time(9, 0, 0)
DAY_END = time(19, 59, 59)



def parse_date(value):
    if isinstance(value, datetime): return value
    if isinstance(value, date): return datetime.combine(value, time.min)
    if isinstance(value, (int, float)):
        # Excel serial date conversion is handled by openpyxl in normal workbook cells.
        return None
    if isinstance(value, str):
        for fmt in ('%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S'):
            try: return datetime.strptime(value.strip(), fmt)
            except ValueError: pass
    return None


def choose_sheet(wb, preferred=None):
    if preferred and preferred in wb.sheetnames: return wb[preferred]
    return wb[wb.sheetnames[0]]


def generate(source_path, template_path, output_path, source_sheet_name='', template_sheet_name='', clear_manual=True, mode='12 годин'):
    src_wb = load_workbook(source_path, data_only=True, read_only=True)
    src = choose_sheet(src_wb, source_sheet_name or None)
    headers = [str(src.cell(1, c).value or '').strip().lower() for c in range(1, src.max_column + 1)]

    # Поточний експорт користувача: A=напрямок (I/E), B=тип транспорту, C=дата пропуску.
    # Колонки визначаються за заголовками; якщо заголовки змінені — використовуємо цей порядок.
    def find_col(terms, fallback):
        for i, h in enumerate(headers, 1):
            if any(term in h for term in terms): return i
        return fallback
    direction_col = find_col(['напрямок'], 1)
    type_col = find_col(['тип транспорту'], 2)
    date_col = find_col(['дата пропуску', 'дата оформлення', 'дата'], 3)

    counts = {}  # (date, shift, category, direction) -> count
    dates_seen = []
    unknown_types = set()
    bad_rows = 0
    for row in src.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < max(direction_col, type_col, date_col): continue
        dt = parse_date(row[date_col-1])
        if dt is None:
            if any(v is not None for v in row): bad_rows += 1
            continue
        direction = str(row[direction_col-1] or '').strip().upper()
        raw_type = row[type_col-1]
        try: typ = int(raw_type)
        except (TypeError, ValueError):
            unknown_types.add(str(raw_type)); continue
        if typ in PASSENGER: category = 'passenger'
        elif typ in FREIGHT: category = 'freight'
        elif typ in BUS: category = 'bus'
        else:
            unknown_types.add(str(typ)); continue
        if direction not in ('E', 'I'):
            continue
        # У 12-годинному звіті ніч прив'язана до дати її завершення:
        # 00:00–08:59 належить поточній даті, а 20:00–23:59 — наступній.
        # Тому для 1-го числа ніч містить лише записи 00:00–08:59 цього числа;
        # записи попереднього місяця не підтягуються.
        if dt.time() >= time(20, 0, 0):
            report_date = dt.date() + __import__('datetime').timedelta(days=1)
            shift = 'ніч'
        elif dt.time() < time(9, 0, 0):
            report_date = dt.date()
            shift = 'ніч'
        else:
            report_date = dt.date()
            shift = 'день'
        key = (report_date, shift, category, direction)
        counts[key] = counts.get(key, 0) + 1
        # Для режиму «доба» календарна дата не зсувається.
        day_key = (dt.date(), 'доба', category, direction)
        counts[day_key] = counts.get(day_key, 0) + 1
        dates_seen.append(dt.date())

    if not dates_seen:
        raise ValueError('Не знайшов жодного рядка з розпізнаною датою. Перевір вихідний файл.')
    report_year, report_month = min(dates_seen).year, min(dates_seen).month
    # If source contains multiple months, choose the month with the most rows and warn in the GUI.
    month_counts = {}
    for d in dates_seen: month_counts[(d.year, d.month)] = month_counts.get((d.year, d.month), 0) + 1
    report_year, report_month = max(month_counts, key=month_counts.get)
    days_in_month = calendar.monthrange(report_year, report_month)[1]

    out_wb = load_workbook(template_path)
    out = choose_sheet(out_wb, template_sheet_name or None)
    # Update report period cell if present.
    for row in out.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and 'з 01.10.2025' in cell.value:
                cell.value = f'з 01.{report_month:02d}.{report_year} по {days_in_month:02d}.{report_month:02d}.{report_year} року'

    # Data starts at row 20. Templates may have either one row per day
    # (режим "доба") or two rows per day (режим "12 годин").
    start_row = 20
    if mode == 'доба':
        expected_last = start_row + days_in_month - 1
    else:
        expected_last = start_row + days_in_month * 2 - 1

    if out.max_row < expected_last:
        raise ValueError(
            f'У шаблоні немає достатньо рядків для режиму «{mode}». '
            f'Очікував рядок {expected_last}, у шаблоні останній рядок {out.max_row}.'
        )

    if mode != 'доба':
        # Detect a daily template (one record row per date). A 12-hour report
        # needs a separate template with vertically merged date cells per pair.
        first_date = out.cell(start_row, 2).value
        second_date = out.cell(start_row + 1, 2).value
        has_first_pair_merge = f'B{start_row}:B{start_row + 1}' in [
            str(rng) for rng in out.merged_cells.ranges
        ]
        if (
            not has_first_pair_merge
            and isinstance(first_date, (datetime, date))
            and isinstance(second_date, (datetime, date))
            and first_date.date() + __import__('datetime').timedelta(days=1)
                == second_date.date()
        ):
            raise ValueError(
                'Обраний Excel-шаблон має режим «доба» (один рядок на дату). '
                'Для режиму «12 годин» обери шаблон із двома рядками на дату.'
            )

    # F12 is the merged-cell anchor in the supplied DОМК-С-1 templates.
    out['F12'] = 'доба' if mode == 'доба' else '12 годин'

    for day_num in range(1, days_in_month + 1):
        day_date = date(report_year, report_month, day_num)

        if mode == 'доба':
            # Daily template: exactly one data row per calendar day.
            # Do not step by two rows: the following rows may be the report footer,
            # which contains merged cells and must remain untouched.
            rr = start_row + day_num - 1
            out.row_dimensions[rr].hidden = False
            out.cell(rr, 1).value = day_num
            out.cell(rr, 2).value = datetime(report_year, report_month, day_num)
            out.cell(rr, 3).value = 'доба'
            rows_to_clear = [rr]
            target_rows = [('доба', rr)]
        else:
            # 12-hour template: two rows per calendar day, with the date
            # merged vertically across the night/day pair.
            first_row = start_row + (day_num - 1) * 2
            second_row = first_row + 1
            out.row_dimensions[first_row].hidden = False
            out.row_dimensions[second_row].hidden = False
            merged_ref = f'B{first_row}:B{second_row}'
            if merged_ref not in [str(rng) for rng in out.merged_cells.ranges]:
                out.merge_cells(merged_ref)
            out.cell(first_row, 1).value = (day_num - 1) * 2 + 1
            out.cell(second_row, 1).value = (day_num - 1) * 2 + 2
            out.cell(first_row, 2).value = datetime(report_year, report_month, day_num)
            out.cell(first_row, 3).value = 'ніч'
            out.cell(second_row, 3).value = 'день'
            rows_to_clear = [first_row, second_row]
            target_rows = [('ніч', first_row), ('день', second_row)]

        for clear_row in rows_to_clear:
            for cc in range(4, 16):
                out.cell(clear_row, cc).value = 0
            if clear_manual:
                out.cell(clear_row, 16).value = None
                out.cell(clear_row, 17).value = None

        for shift, target_row in target_rows:
            for category, cols in {
                'passenger': {'E': 4, 'I': 5},
                'freight': {'E': 8, 'I': 9},
                'bus': {'E': 12, 'I': 13},
            }.items():
                for direction, col in cols.items():
                    out.cell(target_row, col).value = counts.get(
                        (day_date, shift, category, direction), 0
                    )

    # Apply date formatting only to the actual date rows.
    if mode == 'доба':
        date_rows = range(start_row, start_row + days_in_month)
    else:
        date_rows = range(start_row, start_row + days_in_month * 2, 2)
    for r in date_rows:
        out.cell(r, 2).number_format = 'dd.mm.yyyy'

    out_wb.save(output_path)
    src_wb.close()
    return {
        'year': report_year, 'month': report_month, 'days': days_in_month,
        'rows_counted': len(dates_seen), 'unknown_types': sorted(unknown_types), 'bad_rows': bad_rows,
        'source_sheet': src.title, 'template_sheet': out.title,
    }


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Автозаповнення ДОМК-С-1')
        self.geometry('760x390')
        self.minsize(680, 360)
        self.source = tk.StringVar()
        self.template = tk.StringVar()
        self.output = tk.StringVar()
        self.status = tk.StringVar(value='Обери вихідний файл, шаблон і режим звіту.')
        self.mode = tk.StringVar(value='12 годин')
        self._build()

    def _build(self):
        root = ttk.Frame(self, padding=18); root.pack(fill='both', expand=True)
        ttk.Label(root, text='Автозаповнення ДОМК-С-1', font=('Segoe UI', 16, 'bold')).pack(anchor='w')
        ttk.Label(root, text='Обери два Excel-файли. Програма створить окремий заповнений звіт і не змінить оригінали.', wraplength=700).pack(anchor='w', pady=(4, 14))
        self._file_row(root, '1. Вихідні дані Excel:', self.source, self.pick_source)
        self._file_row(root, '2. Шаблон ДОМК-С-1:', self.template, self.pick_template)
        self._file_row(root, '3. Куди зберегти результат:', self.output, self.pick_output)
        mode_frame = ttk.LabelFrame(root, text='Режим звіту', padding=8); mode_frame.pack(fill='x', pady=(8, 2))
        ttk.Radiobutton(mode_frame, text='12 годин (день/ніч)', variable=self.mode, value='12 годин').pack(side='left', padx=(0, 18))
        ttk.Radiobutton(mode_frame, text='Доба (один рядок на дату)', variable=self.mode, value='доба').pack(side='left')
        ttk.Label(root, text='Для режиму «12 годин»: день 09:00–19:59; ніч 20:00–08:59 (ніч 20:00–23:59 відноситься до наступної дати).', foreground='#555', wraplength=700).pack(anchor='w', pady=(4, 2))
        ttk.Label(root, text='Коди: легкові 10/13/20/21; вантажні 14/22/23/24; автобуси 9/29.', foreground='#555').pack(anchor='w')
        bottom = ttk.Frame(root); bottom.pack(fill='x', pady=(16, 0))
        ttk.Button(bottom, text='ЗАПОВНИТИ ТАБЛИЦЮ', command=self.run,).pack(side='left')
        ttk.Label(bottom, textvariable=self.status, wraplength=480).pack(side='left', padx=14)

    def _file_row(self, parent, label, var, command):
        frame = ttk.Frame(parent); frame.pack(fill='x', pady=5)
        ttk.Label(frame, text=label, width=26).pack(side='left')
        ttk.Entry(frame, textvariable=var).pack(side='left', fill='x', expand=True, padx=7)
        ttk.Button(frame, text='Обрати…', command=command).pack(side='right')

    def pick_source(self):
        p = filedialog.askopenfilename(title='Обрати вихідні дані', filetypes=[('Excel', '*.xlsx *.xlsm')])
        if p:
            self.source.set(p)
            if not self.output.get(): self.output.set(os.path.splitext(p)[0] + '_ДОМК_заповнено.xlsx')

    def pick_template(self):
        p = filedialog.askopenfilename(title='Обрати шаблон ДОМК-С-1', filetypes=[('Excel', '*.xlsx *.xlsm')])
        if p: self.template.set(p)

    def pick_output(self):
        p = filedialog.asksaveasfilename(title='Зберегти готовий звіт', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if p: self.output.set(p)

    def run(self):
        if not self.source.get() or not self.template.get() or not self.output.get():
            messagebox.showwarning('Не всі файли обрані', 'Обери вихідний файл, шаблон і місце збереження результату.'); return
        if os.path.abspath(self.output.get()) in (os.path.abspath(self.source.get()), os.path.abspath(self.template.get())):
            messagebox.showerror('Помилка', 'Результат потрібно зберегти в окремий файл, не поверх оригіналу.'); return
        try:
            self.status.set('Обробляю файли…'); self.update_idletasks()
            result = generate(self.source.get(), self.template.get(), self.output.get(), mode=self.mode.get())
            message = (f"Готово!\n\nЗвіт: {result['month']:02d}.{result['year']}\nДнів: {result['days']}\nРежим: {self.mode.get()}\nОброблено записів: {result['rows_counted']}\n")
            if result['unknown_types']:
                message += '\nНевідомі/не враховані коди типів: ' + ', '.join(result['unknown_types'])
            if result['bad_rows']:
                message += f"\nРядків без розпізнаної дати: {result['bad_rows']}"
            message += '\n\nПеревір підсумки перед офіційним використанням.'
            self.status.set('Готово. Файл збережено.')
            messagebox.showinfo('Успішно', message)
        except Exception as e:
            self.status.set('Не вдалося заповнити таблицю.')
            messagebox.showerror('Помилка', f'{e}\n\nДеталі:\n{traceback.format_exc()}')

if __name__ == '__main__':
    App().mainloop()

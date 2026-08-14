import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import subprocess
import sys

class AssetManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт имущества")
        self.root.geometry("1400x750")
        self.root.configure(bg='#f0f0f0')

        self.dark_mode = False
        self.tooltip = None  # для всплывающей подсказки

        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.data_dir = os.path.join(desktop, 'Учёт имущества')
        self.scans_dir = os.path.join(self.data_dir, 'Сканы')
        os.makedirs(self.scans_dir, exist_ok=True)
        self.excel_path = os.path.join(self.data_dir, 'assets.xlsx')

        self.assets = []
        self.filtered_assets = []
        self.grouped_assets = []

        if not os.path.exists(self.excel_path):
            self.create_excel_file()

        self.load_assets()
        self.setup_styles()
        self.setup_ui()
        self.apply_theme()
        self.setup_shortcuts()
        self.auto_save()

    # ---------- Цветовые схемы ----------
    def get_colors(self):
        if self.dark_mode:
            return {
                'bg': '#2e2e2e', 'fg': '#ffffff', 'panel_bg': '#3c3c3c',
                'entry_bg': '#3c3c3c', 'entry_fg': '#ffffff',
                'button_bg': '#555555', 'button_fg': '#ffffff',
                'tree_bg': '#2e2e2e', 'tree_fg': '#ffffff',
                'tree_selected_bg': '#4a90d9', 'tree_selected_fg': '#ffffff',
                'heading_bg': '#1e1e1e', 'heading_fg': '#ffffff',
                'tag_written_off': '#5c3a3a', 'tag_in_repair': '#5c4a2a',
                'filter_bg': '#3c3c3c', 'filter_fg': '#ffffff',
                'stats_bg': '#2e2e2e', 'stats_fg': '#cccccc',
                'detail_bg': '#2e2e2e', 'detail_fg': '#ffffff'
            }
        else:
            return {
                'bg': '#f0f0f0', 'fg': '#333333', 'panel_bg': '#ffffff',
                'entry_bg': '#ffffff', 'entry_fg': '#333333',
                'button_bg': '#4CAF50', 'button_fg': '#ffffff',
                'tree_bg': '#ffffff', 'tree_fg': '#333333',
                'tree_selected_bg': '#2196F3', 'tree_selected_fg': '#ffffff',
                'heading_bg': '#2196F3', 'heading_fg': '#ffffff',
                'tag_written_off': '#FFEBEE', 'tag_in_repair': '#FFF3E0',
                'filter_bg': '#ffffff', 'filter_fg': '#333333',
                'stats_bg': '#f0f0f0', 'stats_fg': '#666666',
                'detail_bg': '#f0f0f0', 'detail_fg': '#333333'
            }

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('default')

    def setup_ui(self):
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Верхняя панель
        self.top_panel = tk.Frame(self.main_container)
        self.top_panel.pack(fill=tk.X, pady=(0, 20))

        self.title_frame = tk.Frame(self.top_panel)
        self.title_frame.pack(side=tk.LEFT)

        self.title_label = tk.Label(self.title_frame, text="📦 Учёт имущества",
                                   font=('Segoe UI', 24, 'bold'))
        self.title_label.pack(side=tk.LEFT)

        self.subtitle_label = tk.Label(self.title_frame, text="by Ольгерд",
                                      font=('Segoe UI', 20))
        self.subtitle_label.pack(side=tk.LEFT, padx=(8, 0))

        self.button_frame = tk.Frame(self.top_panel)
        self.button_frame.pack(side=tk.RIGHT)

        self.add_btn = tk.Button(self.button_frame, text="➕ Добавить",
                                command=self.add_asset, relief=tk.FLAT,
                                cursor='hand2', padx=20, pady=10)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.edit_btn = tk.Button(self.button_frame, text="✏️ Редактировать",
                                 command=self.edit_asset, relief=tk.FLAT,
                                 cursor='hand2', padx=20, pady=10)
        self.edit_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.delete_btn = tk.Button(self.button_frame, text="🗑️ Удалить",
                                   command=self.delete_asset, relief=tk.FLAT,
                                   cursor='hand2', padx=20, pady=10)
        self.delete_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.export_btn = tk.Button(self.button_frame, text="📊 Экспорт",
                                   command=self.export_to_excel, relief=tk.FLAT,
                                   cursor='hand2', padx=20, pady=10)
        self.export_btn.pack(side=tk.LEFT)

        # Панель фильтров
        self.filter_panel = tk.Frame(self.main_container, relief=tk.RAISED, bd=1)
        self.filter_panel.pack(fill=tk.X, pady=(0, 20), padx=2)

        self.inner_filter = tk.Frame(self.filter_panel)
        self.inner_filter.pack(fill=tk.X, padx=10, pady=10)

        self.search_label = tk.Label(self.inner_filter, text="🔍 Поиск:")
        self.search_label.pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_assets())
        self.search_entry = tk.Entry(self.inner_filter, textvariable=self.search_var, width=20)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))

        self.name_label = tk.Label(self.inner_filter, text="Наименование:")
        self.name_label.pack(side=tk.LEFT, padx=(0, 5))
        self.name_filter_var = tk.StringVar(value="Все")
        self.name_combo = ttk.Combobox(self.inner_filter, textvariable=self.name_filter_var,
                                       values=["Все"] + self.get_unique_values('name'),
                                       width=15, state='readonly')
        self.name_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.name_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.location_label = tk.Label(self.inner_filter, text="Расположение:")
        self.location_label.pack(side=tk.LEFT, padx=(0, 5))
        self.location_filter_var = tk.StringVar(value="Все")
        self.location_combo = ttk.Combobox(self.inner_filter, textvariable=self.location_filter_var,
                                           values=["Все"] + self.get_unique_values('location'),
                                           width=15, state='readonly')
        self.location_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.location_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.act_label = tk.Label(self.inner_filter, text="Акт/Накладная:")
        self.act_label.pack(side=tk.LEFT, padx=(0, 5))
        self.act_filter_var = tk.StringVar(value="Все")
        self.act_combo = ttk.Combobox(self.inner_filter, textvariable=self.act_filter_var,
                                      values=["Все"] + self.get_unique_values('act'),
                                      width=15, state='readonly')
        self.act_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.act_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.inv_label = tk.Label(self.inner_filter, text="Инв. номер:")
        self.inv_label.pack(side=tk.LEFT, padx=(0, 5))
        self.inv_filter_var = tk.StringVar(value="Все")
        self.inv_combo = ttk.Combobox(self.inner_filter, textvariable=self.inv_filter_var,
                                      values=["Все"] + self.get_unique_values('inventory'),
                                      width=15, state='readonly')
        self.inv_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.inv_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.direction_label = tk.Label(self.inner_filter, text="Направление:")
        self.direction_label.pack(side=tk.LEFT, padx=(0, 5))
        self.direction_var = tk.StringVar(value="Все")
        self.direction_combo = ttk.Combobox(self.inner_filter, textvariable=self.direction_var,
                                            values=["Все"] + self.get_directions(),
                                            width=15, state='readonly')
        self.direction_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.direction_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.status_label = tk.Label(self.inner_filter, text="Статус:")
        self.status_label.pack(side=tk.LEFT, padx=(0, 5))
        self.status_var = tk.StringVar(value="Все")
        self.status_combo = ttk.Combobox(self.inner_filter, textvariable=self.status_var,
                                         values=["Все", "Активен", "В ремонте", "Списан"],
                                         width=12, state='readonly')
        self.status_combo.pack(side=tk.LEFT)
        self.status_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        # Таблица
        self.table_frame = tk.Frame(self.main_container, relief=tk.SOLID, bd=1)
        self.table_frame.pack(fill=tk.BOTH, expand=True)

        columns = {
            'name': ('Наименование', 300),      # <-- увеличена ширина
            'inventory': ('Инв. номер', 100),
            'qty_unit': ('Кол-во/ед. изм.', 100),
            'direction': ('Направление', 120),
            'location': ('Расположение', 150),
            'issued_to': ('Кому выдано', 150),
            'status': ('Статус', 100),
            'act': ('Акт/Накладная', 120),
            'note': ('Примечание', 200)
        }

        self.tree = ttk.Treeview(self.table_frame, columns=tuple(columns.keys()),
                                 show='headings', selectmode='extended')

        for col, (text, width) in columns.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor='center', stretch=True)

        vscroll = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hscroll = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vscroll.grid(row=0, column=1, sticky='ns')
        hscroll.grid(row=1, column=0, sticky='ew')
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind('<Double-Button-1>', self.on_tree_double_click)
        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)   # для строки деталей
        self.tree.bind('<Motion>', self.on_tree_motion)             # для tooltip
        self.tree.bind('<Leave>', self.hide_tooltip)

        # Строка деталей (под таблицей)
        self.detail_frame = tk.Frame(self.main_container)
        self.detail_frame.pack(fill=tk.X, pady=(5, 0))
        self.detail_label = tk.Label(self.detail_frame, text="", anchor='w',
                                     font=('Segoe UI', 10))
        self.detail_label.pack(fill=tk.X)

        # Нижняя панель со статистикой и кнопкой темы
        self.stats_panel = tk.Frame(self.main_container)
        self.stats_panel.pack(fill=tk.X, pady=(10, 0))
        self.stats_label = tk.Label(self.stats_panel, text="")
        self.stats_label.pack(side=tk.LEFT)

        self.theme_btn = tk.Button(self.stats_panel, text="🌙 Тёмная тема",
                                   command=self.toggle_theme, relief=tk.FLAT,
                                   cursor='hand2', padx=10, pady=5)
        self.theme_btn.pack(side=tk.RIGHT)

        self.update_table()
        self.update_stats()

    def apply_theme(self):
        c = self.get_colors()
        self.root.configure(bg=c['bg'])
        self.main_container.configure(bg=c['bg'])
        self.top_panel.configure(bg=c['bg'])
        self.title_frame.configure(bg=c['bg'])
        self.title_label.configure(bg=c['bg'], fg=c['fg'])
        self.subtitle_label.configure(bg=c['bg'], fg=c['stats_fg'])
        self.button_frame.configure(bg=c['bg'])
        self.filter_panel.configure(bg=c['filter_bg'])
        self.inner_filter.configure(bg=c['filter_bg'])
        self.table_frame.configure(bg=c['filter_bg'])
        self.stats_panel.configure(bg=c['bg'])
        self.stats_label.configure(bg=c['bg'], fg=c['stats_fg'])

        # Новые элементы
        self.detail_frame.configure(bg=c['detail_bg'])
        self.detail_label.configure(bg=c['detail_bg'], fg=c['detail_fg'])

        for btn in [self.add_btn, self.edit_btn, self.delete_btn, self.export_btn, self.theme_btn]:
            btn.configure(bg=c['button_bg'], fg=c['button_fg'],
                         activebackground=c['button_bg'])

        for label in [self.search_label, self.name_label, self.location_label,
                      self.act_label, self.inv_label, self.direction_label, self.status_label]:
            label.configure(bg=c['filter_bg'], fg=c['filter_fg'])

        self.search_entry.configure(bg=c['entry_bg'], fg=c['entry_fg'],
                                    insertbackground=c['entry_fg'])

        self.style.configure('Treeview',
                             background=c['tree_bg'],
                             fieldbackground=c['tree_bg'],
                             foreground=c['tree_fg'],
                             rowheight=30,
                             font=('Segoe UI', 10))
        self.style.configure('Treeview.Heading',
                             background=c['heading_bg'],
                             foreground=c['heading_fg'],
                             font=('Segoe UI', 10, 'bold'),
                             relief='flat')
        self.style.map('Treeview',
                       background=[('selected', c['tree_selected_bg'])],
                       foreground=[('selected', c['tree_selected_fg'])])
        self.style.map('Treeview.Heading',
                       background=[('active', c['heading_bg'])])
        self.style.configure('TCombobox',
                             fieldbackground=c['entry_bg'],
                             background=c['entry_bg'],
                             foreground=c['entry_fg'],
                             arrowcolor=c['entry_fg'])

        self.tree.tag_configure('written_off', background=c['tag_written_off'])
        self.tree.tag_configure('in_repair', background=c['tag_in_repair'])
        self.tree.tag_configure('group', font=('Segoe UI', 10, 'bold'))

        self.theme_btn.config(text="☀️ Светлая тема" if self.dark_mode else "🌙 Тёмная тема")

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    # ---------- Tooltip ----------
    def on_tree_motion(self, event):
        region = self.tree.identify('region', event.x, event.y)
        if region == 'cell':
            col = self.tree.identify_column(event.x)
            row = self.tree.identify_row(event.y)
            if row and col:
                values = self.tree.item(row, 'values')
                col_index = int(col.replace('#', '')) - 1
                if 0 <= col_index < len(values):
                    text = str(values[col_index])
                    if len(text) > 20:  # порог, можно настроить
                        self.show_tooltip(event.x_root, event.y_root, text)
                    else:
                        self.hide_tooltip()
                else:
                    self.hide_tooltip()
            else:
                self.hide_tooltip()
        else:
            self.hide_tooltip()

    def show_tooltip(self, x, y, text):
        self.hide_tooltip()
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x+10}+{y+10}")
        # Цвет tooltip всегда жёлтый, текст чёрный для читаемости
        label = tk.Label(self.tooltip, text=text, background='#ffffe0',
                         relief='solid', borderwidth=1, padx=5, pady=2,
                         wraplength=400, justify='left', fg='#000000')
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    # ---------- Строка деталей ----------
    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            self.detail_label.config(text="")
            return
        item = selected[0]
        values = self.tree.item(item, 'values')
        children = self.tree.get_children(item)
        if children:
            # Группа
            name = values[0] if values else ""
            qty = values[2] if len(values) > 2 else ""
            self.detail_label.config(text=f"Группа: {name} | Всего: {qty}")
        else:
            # Конкретный актив
            if values:
                name = values[0] if values[0] else "—"
                inv = values[1] if values[1] else "—"
                self.detail_label.config(text=f"Актив: {name} | Инв. номер: {inv}")
            else:
                self.detail_label.config(text="")

    # ---------- Остальные методы (без изменений) ----------
    def get_unique_values(self, field):
        values = set()
        for asset in self.assets:
            val = asset.get(field, '')
            if val:
                values.add(str(val))
        return sorted(list(values))

    def get_directions(self):
        return self.get_unique_values('direction')

    def update_filter_values(self):
        self.name_combo['values'] = ["Все"] + self.get_unique_values('name')
        self.location_combo['values'] = ["Все"] + self.get_unique_values('location')
        self.act_combo['values'] = ["Все"] + self.get_unique_values('act')
        self.inv_combo['values'] = ["Все"] + self.get_unique_values('inventory')
        self.direction_combo['values'] = ["Все"] + self.get_directions()

    def create_excel_file(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Активы"
        headers = [
            'Наименование', 'Инв. номер', 'Количество', 'Ед. измерения',
            'Направление', 'Расположение', 'Кому выдано', 'Статус',
            'Акт/Накладная', 'Примечание'
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF', size=12)
            cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        widths = [25, 15, 10, 12, 20, 20, 20, 12, 20, 25]
        for col, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        wb.save(self.excel_path)

    def load_assets(self):
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            self.assets = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    asset = {
                        'name': row[0] or '',
                        'inventory': row[1] or '',
                        'quantity': row[2] if isinstance(row[2], (int, float)) else 1,
                        'unit': row[3] or 'шт.',
                        'direction': row[4] or '',
                        'location': row[5] or '',
                        'issued_to': row[6] or '',
                        'status': row[7] or 'Активен',
                        'act': row[8] or '',
                        'note': row[9] or ''
                    }
                    self.assets.append(asset)
            self.filtered_assets = self.assets.copy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")

    def save_to_excel(self):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Активы"
            headers = [
                'Наименование', 'Инв. номер', 'Количество', 'Ед. измерения',
                'Направление', 'Расположение', 'Кому выдано', 'Статус',
                'Акт/Накладная', 'Примечание'
            ]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF', size=12)
                cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for row_idx, asset in enumerate(self.assets, 2):
                ws.cell(row=row_idx, column=1, value=asset['name'])
                ws.cell(row=row_idx, column=2, value=asset['inventory'])
                ws.cell(row=row_idx, column=3, value=asset['quantity'])
                ws.cell(row=row_idx, column=4, value=asset['unit'])
                ws.cell(row=row_idx, column=5, value=asset['direction'])
                ws.cell(row=row_idx, column=6, value=asset['location'])
                ws.cell(row=row_idx, column=7, value=asset['issued_to'])
                ws.cell(row=row_idx, column=8, value=asset['status'])
                ws.cell(row=row_idx, column=9, value=asset['act'])
                ws.cell(row=row_idx, column=10, value=asset['note'])
                if row_idx % 2 == 0:
                    for col in range(1, 11):
                        ws.cell(row=row_idx, column=col).fill = PatternFill(
                            start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
            widths = [25, 15, 10, 12, 20, 20, 20, 12, 20, 25]
            for col, width in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = width
            ws.auto_filter.ref = f"A1:J{len(self.assets) + 1}"
            ws.freeze_panes = 'A2'
            wb.save(self.excel_path)
            return True
        except PermissionError:
            # Если файл заблокирован, пробуем сохранить во временную папку
            backup_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Temp', 'assets_backup.xlsx')
            try:
                wb.save(backup_path)
                messagebox.showwarning(
                    "Внимание",
                    f"Файл {self.excel_path} занят другим процессом (возможно, открыт в Excel).\n"
                    f"Данные сохранены во временный файл:\n{backup_path}\n\n"
                    "Закройте Excel и повторите сохранение, либо скопируйте временный файл вручную."
                )
            except Exception as e2:
                messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {str(e2)}")
            return False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {str(e)}")
            return False

    def build_grouped_assets(self):
        groups = {}
        for asset in self.filtered_assets:
            key = asset['name'].strip().lower()
            if key not in groups:
                groups[key] = {
                    'name': asset['name'],
                    'assets': [],
                    'total_qty': 0,
                    'unit': asset['unit'],
                    'inventories': []
                }
            groups[key]['assets'].append(asset)
            groups[key]['total_qty'] += asset['quantity']
            groups[key]['inventories'].append(asset['inventory'])
        self.grouped_assets = list(groups.values())

    def on_tree_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        if region == 'cell':
            col = self.tree.identify_column(event.x)
            if col == '#8':  # Акт/Накладная
                item = self.tree.identify_row(event.y)
                if item:
                    values = self.tree.item(item, 'values')
                    if values and len(values) >= 8:
                        act = values[7]
                        self.open_scan(act)

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        children = self.tree.get_children(item)
        if children:
            current_state = self.tree.item(item, 'open')
            self.tree.item(item, open=not current_state)
            values = list(self.tree.item(item, 'values'))
            if values:
                if not current_state:
                    values[0] = values[0].replace('▸ ', '▾ ')
                else:
                    values[0] = values[0].replace('▾ ', '▸ ')
                self.tree.item(item, values=values)
            return "break"
        else:
            self.edit_asset()
            return "break"

    def open_scan(self, act_name):
        if not act_name:
            return
        target = None
        for f in os.listdir(self.scans_dir):
            if act_name.lower() in f.lower():
                target = os.path.join(self.scans_dir, f)
                break
        if target and os.path.exists(target):
            try:
                if sys.platform == 'win32':
                    os.startfile(target)
                elif sys.platform == 'darwin':
                    subprocess.call(['open', target])
                else:
                    subprocess.call(['xdg-open', target])
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")
        else:
            messagebox.showinfo("Информация", f"Скан не найден. Разместите файл в папке:\n{self.scans_dir}\n\nИскомое название: {act_name}")

    def add_asset(self):
        dialog = AssetDialog(self.root, "Добавить актив")
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            if not dialog.result.get('inventory'):
                dialog.result['inventory'] = self.generate_inventory_number()
            self.assets.append(dialog.result)
            self.save_to_excel()
            self.refresh_after_change()
            messagebox.showinfo("Успех", "Актив успешно добавлен")

    def edit_asset(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите актив для редактирования")
            return
        item = selected[0]
        children = self.tree.get_children(item)
        if children:
            messagebox.showinfo("Информация", "Выберите конкретный актив внутри группы")
            return
        values = self.tree.item(item, 'values')
        if not values or len(values) < 2:
            messagebox.showerror("Ошибка", "Не удалось определить актив")
            return
        inventory = values[1]
        asset_index = None
        for i, asset in enumerate(self.assets):
            if str(asset['inventory']) == str(inventory):
                asset_index = i
                break
        if asset_index is None:
            messagebox.showerror("Ошибка", "Актив не найден в базе данных")
            return
        dialog = AssetDialog(self.root, "Редактировать актив", self.assets[asset_index])
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            self.assets[asset_index] = dialog.result
            self.save_to_excel()
            self.refresh_after_change()
            messagebox.showinfo("Успех", "Актив успешно обновлён")

    def delete_asset(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите активы для удаления")
            return
        if not messagebox.askyesno("Подтверждение",
                                   f"Удалить выбранные активы ({len(selected)} шт.)?"):
            return
        inv_numbers = []
        for item in selected:
            children = self.tree.get_children(item)
            if children:
                for child in children:
                    vals = self.tree.item(child, 'values')
                    if vals and len(vals) >= 2:
                        inv_numbers.append(vals[1])
            else:
                vals = self.tree.item(item, 'values')
                if vals and len(vals) >= 2:
                    inv_numbers.append(vals[1])
        if not inv_numbers:
            return
        self.assets = [a for a in self.assets if str(a['inventory']) not in inv_numbers]
        self.save_to_excel()
        self.refresh_after_change()
        messagebox.showinfo("Успех", "Активы успешно удалены")

    def generate_inventory_number(self):
        if not self.assets:
            return "INV-0001"
        max_num = 0
        for asset in self.assets:
            inv = str(asset['inventory'])
            if inv.startswith("INV-"):
                try:
                    num = int(inv.split("-")[1])
                    max_num = max(max_num, num)
                except:
                    pass
        return f"INV-{max_num + 1:04d}"

    def refresh_after_change(self):
        self.update_filter_values()
        self.filter_assets()

    def filter_assets(self):
        search_text = self.search_var.get().lower()
        self.filtered_assets = []
        for asset in self.assets:
            if self.name_filter_var.get() != "Все" and asset['name'] != self.name_filter_var.get():
                continue
            if self.location_filter_var.get() != "Все" and asset['location'] != self.location_filter_var.get():
                continue
            if self.act_filter_var.get() != "Все" and asset['act'] != self.act_filter_var.get():
                continue
            if self.inv_filter_var.get() != "Все" and str(asset['inventory']) != self.inv_filter_var.get():
                continue
            if self.direction_var.get() != "Все" and asset['direction'] != self.direction_var.get():
                continue
            if self.status_var.get() != "Все" and asset['status'] != self.status_var.get():
                continue
            if search_text:
                searchable = ' '.join([
                    str(asset.get('name', '')),
                    str(asset.get('inventory', '')),
                    str(asset.get('location', '')),
                    str(asset.get('direction', '')),
                    str(asset.get('act', '')),
                    str(asset.get('issued_to', '')),
                    str(asset.get('note', ''))
                ]).lower()
                if search_text not in searchable:
                    continue
            self.filtered_assets.append(asset)
        self.build_grouped_assets()
        self.update_table()
        self.update_stats()

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for group in self.grouped_assets:
            if len(group['assets']) == 1:
                asset = group['assets'][0]
                tags = self.get_tags(asset)
                self.tree.insert('', 'end', values=(
                    asset['name'],
                    asset['inventory'],
                    f"{asset['quantity']}/{asset['unit']}",
                    asset['direction'],
                    asset['location'],
                    asset['issued_to'],
                    asset['status'],
                    asset['act'],
                    asset['note']
                ), tags=tags)
            else:
                group_tags = ('group',)
                group_name = '▸ ' + group['name']
                group_id = self.tree.insert('', 'end', values=(
                    group_name,
                    '',
                    f"{group['total_qty']}/{group['unit']}",
                    '',
                    '',
                    '',
                    '',
                    '',
                    ''
                ), tags=group_tags, open=False)
                for asset in group['assets']:
                    tags = self.get_tags(asset)
                    self.tree.insert(group_id, 'end', values=(
                        '',
                        asset['inventory'],
                        f"{asset['quantity']}/{asset['unit']}",
                        asset['direction'],
                        asset['location'],
                        asset['issued_to'],
                        asset['status'],
                        asset['act'],
                        asset['note']
                    ), tags=tags)

    def get_tags(self, asset):
        tags = []
        if asset['status'] == 'Списан':
            tags.append('written_off')
        elif asset['status'] == 'В ремонте':
            tags.append('in_repair')
        return tags

    def update_stats(self):
        total = len(self.assets)
        active = len([a for a in self.assets if a['status'] == 'Активен'])
        in_repair = len([a for a in self.assets if a['status'] == 'В ремонте'])
        written_off = len([a for a in self.assets if a['status'] == 'Списан'])
        self.stats_label.config(
            text=f"Всего: {total} | Активных: {active} | В ремонте: {in_repair} | "
                 f"Списано: {written_off}"
        )

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="assets_export.xlsx"
        )
        if file_path:
            try:
                import shutil
                shutil.copy2(self.excel_path, file_path)
                messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось экспортировать: {str(e)}")

    def auto_save(self):
        self.save_to_excel()
        self.root.after(300000, self.auto_save)

    def on_closing(self):
        if messagebox.askyesno("Выход", "Сохранить изменения перед выходом?"):
            if self.save_to_excel():
                self.root.destroy()
        else:
            self.root.destroy()

    def setup_shortcuts(self):
        self.root.bind('<Control-n>', lambda e: self.add_asset())
        self.root.bind('<Control-e>', lambda e: self.edit_asset())
        self.root.bind('<Delete>', lambda e: self.delete_asset())
        self.root.bind('<Control-f>', lambda e: self.focus_search())
        self.root.bind('<Control-s>', lambda e: self.save_to_excel())

    def focus_search(self):
        self.search_entry.focus_set()


class AssetDialog:
    def __init__(self, parent, title, asset=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("550x650")
        self.dialog.configure(bg='#f0f0f0')
        self.dialog.resizable(False, False)

        self.result = None
        self.asset = asset or {}

        self.setup_ui()

        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.focus_force()
        self.dialog.lift()

    def setup_ui(self):
        tk.Label(self.dialog, text="Информация об активе",
                font=('Segoe UI', 16, 'bold'),
                bg='#f0f0f0', fg='#333333').pack(pady=20)

        form_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        form_frame.pack(fill=tk.BOTH, expand=True, padx=30)

        fields = [
            ('Наименование:', 'name', True),
            ('Инвентарный номер:', 'inventory', False),
            ('Количество:', 'quantity', False),
            ('Ед. измерения:', 'unit', False),
            ('Направление:', 'direction', True),
            ('Расположение:', 'location', True),
            ('Кому выдано:', 'issued_to', False),
            ('Статус:', 'status', False),
            ('Акт/Накладная:', 'act', False),
            ('Примечание:', 'note', False)
        ]

        self.entries = {}

        for i, (label, key, required) in enumerate(fields):
            tk.Label(form_frame, text=label,
                    bg='#f0f0f0', font=('Segoe UI', 10)).grid(row=i, column=0, sticky='w', pady=5)

            if key == 'status':
                var = tk.StringVar(value=self.asset.get(key, 'Активен'))
                entry = ttk.Combobox(form_frame, textvariable=var,
                                    values=['Активен', 'В ремонте', 'Списан'],
                                    state='readonly', width=25)
            else:
                default = self.asset.get(key, '')
                if key == 'quantity':
                    default = self.asset.get(key, 1)
                elif key == 'unit':
                    default = self.asset.get(key, 'шт.')
                var = tk.StringVar(value=default)
                entry = tk.Entry(form_frame, textvariable=var, width=27,
                               font=('Segoe UI', 10))

            entry.grid(row=i, column=1, sticky='w', pady=5, padx=(10, 0))
            self.entries[key] = var

            if required:
                tk.Label(form_frame, text="*", fg='red',
                        bg='#f0f0f0', font=('Segoe UI', 10, 'bold')).grid(row=i, column=2, sticky='w')

        button_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="💾 Сохранить",
                 command=self.save,
                 bg='#4CAF50', fg='white',
                 font=('Segoe UI', 10, 'bold'),
                 relief=tk.FLAT, cursor='hand2',
                 padx=20, pady=10).pack(side=tk.LEFT, padx=10)

        tk.Button(button_frame, text="Отмена",
                 command=self.dialog.destroy,
                 bg='#9E9E9E', fg='white',
                 font=('Segoe UI', 10),
                 relief=tk.FLAT, cursor='hand2',
                 padx=20, pady=10).pack(side=tk.LEFT, padx=10)

    def save(self):
        required_fields = ['name', 'direction', 'location']
        for field in required_fields:
            if not self.entries[field].get().strip():
                messagebox.showwarning("Внимание",
                                      f"Поле '{field}' обязательно для заполнения")
                return

        try:
            quantity = int(self.entries['quantity'].get().strip()) if self.entries['quantity'].get().strip() else 1
        except ValueError:
            messagebox.showerror("Ошибка", "Количество должно быть числом")
            return

        self.result = {
            'name': self.entries['name'].get().strip(),
            'inventory': self.entries['inventory'].get().strip(),
            'quantity': quantity,
            'unit': self.entries['unit'].get().strip() or 'шт.',
            'direction': self.entries['direction'].get().strip(),
            'location': self.entries['location'].get().strip(),
            'issued_to': self.entries['issued_to'].get().strip(),
            'status': self.entries['status'].get(),
            'act': self.entries['act'].get().strip(),
            'note': self.entries['note'].get().strip()
        }

        self.dialog.destroy()


def main():
    root = tk.Tk()
    app = AssetManager(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()

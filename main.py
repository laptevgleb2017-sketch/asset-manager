import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

class AssetManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт имущества")
        self.root.geometry("1400x750")
        self.root.configure(bg='#f0f0f0')

        # Переменная темы
        self.dark_mode = False

        # Папка для хранения данных на рабочем столе
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        data_dir = os.path.join(desktop, 'Учёт имущества')
        os.makedirs(data_dir, exist_ok=True)
        self.excel_path = os.path.join(data_dir, 'assets.xlsx')

        self.assets = []
        self.filtered_assets = []

        if not os.path.exists(self.excel_path):
            self.create_excel_file()

        self.load_assets()
        self.setup_styles()
        self.setup_ui()
        self.apply_theme()          # применяем светлую тему по умолчанию
        self.setup_shortcuts()
        self.auto_save()

    # ---------- Цветовые схемы ----------
    def get_colors(self):
        if self.dark_mode:
            return {
                'bg': '#2e2e2e',
                'fg': '#ffffff',
                'panel_bg': '#3c3c3c',
                'entry_bg': '#3c3c3c',
                'entry_fg': '#ffffff',
                'button_bg': '#555555',
                'button_fg': '#ffffff',
                'tree_bg': '#2e2e2e',
                'tree_fg': '#ffffff',
                'tree_selected_bg': '#4a90d9',
                'tree_selected_fg': '#ffffff',
                'heading_bg': '#1e1e1e',
                'heading_fg': '#ffffff',
                'tag_written_off': '#5c3a3a',
                'tag_in_repair': '#5c4a2a',
                'filter_bg': '#3c3c3c',
                'filter_fg': '#ffffff',
                'stats_bg': '#2e2e2e',
                'stats_fg': '#cccccc'
            }
        else:
            return {
                'bg': '#f0f0f0',
                'fg': '#333333',
                'panel_bg': '#ffffff',
                'entry_bg': '#ffffff',
                'entry_fg': '#333333',
                'button_bg': '#4CAF50',
                'button_fg': '#ffffff',
                'tree_bg': '#ffffff',
                'tree_fg': '#333333',
                'tree_selected_bg': '#2196F3',
                'tree_selected_fg': '#ffffff',
                'heading_bg': '#2196F3',
                'heading_fg': '#ffffff',
                'tag_written_off': '#FFEBEE',
                'tag_in_repair': '#FFF3E0',
                'filter_bg': '#ffffff',
                'filter_fg': '#333333',
                'stats_bg': '#f0f0f0',
                'stats_fg': '#666666'
            }

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('default')   # тема с видимой сеткой

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

        # Панель поиска и фильтрации
        self.filter_panel = tk.Frame(self.main_container, relief=tk.RAISED, bd=1)
        self.filter_panel.pack(fill=tk.X, pady=(0, 20), padx=2)

        self.inner_filter = tk.Frame(self.filter_panel)
        self.inner_filter.pack(fill=tk.X, padx=10, pady=10)

        self.search_label = tk.Label(self.inner_filter, text="🔍 Поиск:")
        self.search_label.pack(side=tk.LEFT, padx=(0, 5))

        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_assets())
        self.search_entry = tk.Entry(self.inner_filter, textvariable=self.search_var,
                                    width=30)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 20))

        self.direction_label = tk.Label(self.inner_filter, text="Направление:")
        self.direction_label.pack(side=tk.LEFT, padx=(0, 5))

        self.direction_var = tk.StringVar(value="Все")
        self.direction_combo = ttk.Combobox(self.inner_filter, textvariable=self.direction_var,
                                            values=["Все"] + self.get_directions(),
                                            width=15, state='readonly')
        self.direction_combo.pack(side=tk.LEFT, padx=(0, 20))
        self.direction_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        self.status_label = tk.Label(self.inner_filter, text="Статус:")
        self.status_label.pack(side=tk.LEFT, padx=(0, 5))

        self.status_var = tk.StringVar(value="Все")
        self.status_combo = ttk.Combobox(self.inner_filter, textvariable=self.status_var,
                                         values=["Все", "Активен", "В ремонте", "Списан"],
                                         width=15, state='readonly')
        self.status_combo.pack(side=tk.LEFT)
        self.status_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        # Таблица
        self.table_frame = tk.Frame(self.main_container, relief=tk.SOLID, bd=1)
        self.table_frame.pack(fill=tk.BOTH, expand=True)

        columns = {
            'inventory': ('Инв. номер', 100),
            'name': ('Наименование', 200),
            'quantity': ('Кол-во', 70),
            'unit': ('Ед. изм.', 80),
            'direction': ('Направление', 120),
            'location': ('Расположение', 150),
            'responsible': ('Ответственный', 150),
            'cost': ('Стоимость', 100),
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

        self.tree.bind('<Double-Button-1>', lambda e: self.edit_asset())

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

        # Основные контейнеры
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

        # Кнопки действий
        self.add_btn.configure(bg='#4CAF50' if not self.dark_mode else '#2d6a2e',
                               fg='white', activebackground='#45a049' if not self.dark_mode else '#3d8a3e')
        self.edit_btn.configure(bg='#FF9800' if not self.dark_mode else '#b36b00',
                                fg='white', activebackground='#fb8c00' if not self.dark_mode else '#c47d00')
        self.delete_btn.configure(bg='#F44336' if not self.dark_mode else '#a93226',
                                  fg='white', activebackground='#d32f2f' if not self.dark_mode else '#b03a2e')
        self.export_btn.configure(bg='#2196F3' if not self.dark_mode else '#1565c0',
                                  fg='white', activebackground='#1e88e5' if not self.dark_mode else '#1976d2')
        self.theme_btn.configure(bg=c['button_bg'], fg=c['button_fg'],
                                 activebackground=c['button_bg'])

        # Метки фильтра
        self.search_label.configure(bg=c['filter_bg'], fg=c['filter_fg'])
        self.direction_label.configure(bg=c['filter_bg'], fg=c['filter_fg'])
        self.status_label.configure(bg=c['filter_bg'], fg=c['filter_fg'])

        # Поле поиска
        self.search_entry.configure(bg=c['entry_bg'], fg=c['entry_fg'],
                                    insertbackground=c['entry_fg'])

        # Настройка ttk стилей
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

        # Комбобоксы
        self.style.configure('TCombobox',
                             fieldbackground=c['entry_bg'],
                             background=c['entry_bg'],
                             foreground=c['entry_fg'],
                             arrowcolor=c['entry_fg'])

        # Обновляем теги строк
        self.tree.tag_configure('written_off', background=c['tag_written_off'])
        self.tree.tag_configure('in_repair', background=c['tag_in_repair'])

        # Обновить текст кнопки темы
        self.theme_btn.config(text="☀️ Светлая тема" if self.dark_mode else "🌙 Тёмная тема")

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    # ----- Остальные методы (как в предыдущих версиях) -----
    def setup_shortcuts(self):
        self.root.bind('<Control-n>', lambda e: self.add_asset())
        self.root.bind('<Control-e>', lambda e: self.edit_asset())
        self.root.bind('<Delete>', lambda e: self.delete_asset())
        self.root.bind('<Control-f>', lambda e: self.focus_search())
        self.root.bind('<Control-s>', lambda e: self.save_to_excel())

    def focus_search(self):
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.focus_set()
                break

    def create_excel_file(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Активы"

        headers = [
            'Инв. номер', 'Наименование', 'Количество', 'Ед. измерения',
            'Направление', 'Расположение', 'Ответственный', 'Стоимость',
            'Статус', 'Акт/Накладная', 'Примечание'
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF', size=12)
            cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(bottom=Side(style='thin'))

        widths = [15, 25, 10, 12, 20, 20, 20, 12, 12, 20, 25]
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
                        'inventory': row[0],
                        'name': row[1] or '',
                        'quantity': row[2] if isinstance(row[2], (int, float)) else 1,
                        'unit': row[3] or 'шт.',
                        'direction': row[4] or '',
                        'location': row[5] or '',
                        'responsible': row[6] or '',
                        'cost': row[7] if isinstance(row[7], (int, float)) else 0,
                        'status': row[8] or 'Активен',
                        'act': row[9] or '',
                        'note': row[10] or ''
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
                'Инв. номер', 'Наименование', 'Количество', 'Ед. измерения',
                'Направление', 'Расположение', 'Ответственный', 'Стоимость',
                'Статус', 'Акт/Накладная', 'Примечание'
            ]

            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF', size=12)
                cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')

            for row_idx, asset in enumerate(self.assets, 2):
                ws.cell(row=row_idx, column=1, value=asset['inventory'])
                ws.cell(row=row_idx, column=2, value=asset['name'])
                ws.cell(row=row_idx, column=3, value=asset['quantity'])
                ws.cell(row=row_idx, column=4, value=asset['unit'])
                ws.cell(row=row_idx, column=5, value=asset['direction'])
                ws.cell(row=row_idx, column=6, value=asset['location'])
                ws.cell(row=row_idx, column=7, value=asset['responsible'])
                ws.cell(row=row_idx, column=8, value=asset['cost'])
                ws.cell(row=row_idx, column=9, value=asset['status'])
                ws.cell(row=row_idx, column=10, value=asset['act'])
                ws.cell(row=row_idx, column=11, value=asset['note'])

                if row_idx % 2 == 0:
                    for col in range(1, 12):
                        ws.cell(row=row_idx, column=col).fill = PatternFill(
                            start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')

            widths = [15, 25, 10, 12, 20, 20, 20, 12, 12, 20, 25]
            for col, width in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = width

            ws.auto_filter.ref = f"A1:K{len(self.assets) + 1}"
            ws.freeze_panes = 'A2'

            wb.save(self.excel_path)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {str(e)}")
            return False

    def add_asset(self):
        dialog = AssetDialog(self.root, "Добавить актив")
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            if not dialog.result.get('inventory'):
                dialog.result['inventory'] = self.generate_inventory_number()
            self.assets.append(dialog.result)
            self.save_to_excel()
            self.update_filter_values()
            self.filter_assets()
            messagebox.showinfo("Успех", "Актив успешно добавлен")

    def edit_asset(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите актив для редактирования")
            return

        item = self.tree.item(selected[0])
        inventory = item['values'][0]

        asset_index = None
        for i, asset in enumerate(self.assets):
            if asset['inventory'] == inventory:
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
            self.update_filter_values()
            self.filter_assets()
            messagebox.showinfo("Успех", "Актив успешно обновлён")

    def delete_asset(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите активы для удаления")
            return

        if messagebox.askyesno("Подтверждение",
                              f"Удалить выбранные активы ({len(selected)} шт.)?"):
            inventories = []
            for item in selected:
                values = self.tree.item(item)['values']
                inventories.append(values[0])

            self.assets = [a for a in self.assets if a['inventory'] not in inventories]
            self.save_to_excel()
            self.update_filter_values()
            self.filter_assets()
            messagebox.showinfo("Успех", "Активы успешно удалены")

    def generate_inventory_number(self):
        if not self.assets:
            return "INV-0001"
        max_num = 0
        for asset in self.assets:
            inv = asset['inventory']
            if inv.startswith("INV-"):
                try:
                    num = int(inv.split("-")[1])
                    max_num = max(max_num, num)
                except:
                    pass
        return f"INV-{max_num + 1:04d}"

    def get_directions(self):
        directions = set()
        for asset in self.assets:
            if asset['direction']:
                directions.add(asset['direction'])
        return sorted(list(directions))

    def update_filter_values(self):
        current = self.direction_var.get()
        values = ["Все"] + self.get_directions()
        self.direction_combo['values'] = values
        if current not in values:
            self.direction_var.set("Все")

    def filter_assets(self):
        search_text = self.search_var.get().lower()
        direction = self.direction_var.get()
        status = self.status_var.get()

        self.filtered_assets = []
        for asset in self.assets:
            if search_text:
                searchable = f"{asset['name']} {asset['inventory']} {asset['location']} {asset['responsible']} {asset['direction']}"
                if search_text not in searchable.lower():
                    continue

            if direction != "Все" and asset['direction'] != direction:
                continue

            if status != "Все" and asset['status'] != status:
                continue

            self.filtered_assets.append(asset)

        self.update_table()
        self.update_stats()

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for asset in self.filtered_assets:
            tags = []
            if asset['status'] == 'Списан':
                tags.append('written_off')
            elif asset['status'] == 'В ремонте':
                tags.append('in_repair')

            self.tree.insert('', 'end', values=(
                asset['inventory'],
                asset['name'],
                asset['quantity'],
                asset['unit'],
                asset['direction'],
                asset['location'],
                asset['responsible'],
                asset['cost'],
                asset['status'],
                asset['act'],
                asset['note']
            ), tags=tags)

        self.tree.tag_configure('written_off', background='#FFEBEE')
        self.tree.tag_configure('in_repair', background='#FFF3E0')

    def update_stats(self):
        total = len(self.assets)
        active = len([a for a in self.assets if a['status'] == 'Активен'])
        in_repair = len([a for a in self.assets if a['status'] == 'В ремонте'])
        written_off = len([a for a in self.assets if a['status'] == 'Списан'])
        total_cost = sum(a['cost'] for a in self.assets if isinstance(a['cost'], (int, float)))
        self.stats_label.config(
            text=f"Всего: {total} | Активных: {active} | В ремонте: {in_repair} | "
                 f"Списано: {written_off} | Общая стоимость: {total_cost:,.2f} ₽"
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
            ('Инвентарный номер:', 'inventory', False),
            ('Наименование:', 'name', True),
            ('Количество:', 'quantity', False),
            ('Ед. измерения:', 'unit', False),
            ('Направление:', 'direction', True),
            ('Расположение:', 'location', True),
            ('Ответственный:', 'responsible', True),
            ('Стоимость:', 'cost', False),
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
                elif key == 'cost':
                    default = self.asset.get(key, 0)

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
        required_fields = ['name', 'direction', 'location', 'responsible']
        for field in required_fields:
            if not self.entries[field].get().strip():
                messagebox.showwarning("Внимание",
                                      f"Поле '{field}' обязательно для заполнения")
                return

        try:
            quantity = int(self.entries['quantity'].get().strip()) if self.entries['quantity'].get().strip() else 1
            cost = float(self.entries['cost'].get().replace(' ', '').replace(',', '.')) if self.entries['cost'].get().strip() else 0
        except ValueError:
            messagebox.showerror("Ошибка", "Количество и стоимость должны быть числами")
            return

        self.result = {
            'inventory': self.entries['inventory'].get().strip(),
            'name': self.entries['name'].get().strip(),
            'quantity': quantity,
            'unit': self.entries['unit'].get().strip() or 'шт.',
            'direction': self.entries['direction'].get().strip(),
            'location': self.entries['location'].get().strip(),
            'responsible': self.entries['responsible'].get().strip(),
            'cost': cost,
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

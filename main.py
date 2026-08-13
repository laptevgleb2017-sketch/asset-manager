import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

class AssetManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт имущества")
        self.root.geometry("1200x700")
        self.root.configure(bg='#f0f0f0')

        # Путь к Excel файлу
        self.excel_path = "assets.xlsx"
        self.assets = []
        self.filtered_assets = []

        # Создание Excel файла если не существует
        if not os.path.exists(self.excel_path):
            self.create_excel_file()

        # Загрузка данных
        self.load_assets()

        # Настройка стилей
        self.setup_styles()

        # Создание интерфейса
        self.setup_ui()

        # Привязка горячих клавиш
        self.setup_shortcuts()

        # Автосохранение
        self.auto_save()

    def setup_styles(self):
        """Настройка стилей интерфейса"""
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('Treeview',
                       background='white',
                       fieldbackground='white',
                       foreground='#333333',
                       rowheight=35,
                       font=('Segoe UI', 10))

        style.configure('Treeview.Heading',
                       background='#2196F3',
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'),
                       relief='flat')

        style.map('Treeview',
                 background=[('selected', '#2196F3')],
                 foreground=[('selected', 'white')])

        style.map('Treeview.Heading',
                 background=[('active', '#1976D2')])

        style.configure('Action.TButton',
                       font=('Segoe UI', 10),
                       padding=10)

    def setup_ui(self):
        """Создание пользовательского интерфейса"""
        main_container = tk.Frame(self.root, bg='#f0f0f0')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Верхняя панель
        top_panel = tk.Frame(main_container, bg='#f0f0f0')
        top_panel.pack(fill=tk.X, pady=(0, 20))

        title_label = tk.Label(top_panel, text="📦 Учёт имущества",
                              font=('Segoe UI', 24, 'bold'),
                              bg='#f0f0f0', fg='#333333')
        title_label.pack(side=tk.LEFT)

        button_frame = tk.Frame(top_panel, bg='#f0f0f0')
        button_frame.pack(side=tk.RIGHT)

        add_btn = tk.Button(button_frame, text="➕ Добавить",
                           command=self.add_asset,
                           bg='#4CAF50', fg='white',
                           font=('Segoe UI', 10, 'bold'),
                           relief=tk.FLAT, cursor='hand2',
                           padx=20, pady=10)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))

        edit_btn = tk.Button(button_frame, text="✏️ Редактировать",
                            command=self.edit_asset,
                            bg='#FF9800', fg='white',
                            font=('Segoe UI', 10),
                            relief=tk.FLAT, cursor='hand2',
                            padx=20, pady=10)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))

        delete_btn = tk.Button(button_frame, text="🗑️ Удалить",
                              command=self.delete_asset,
                              bg='#F44336', fg='white',
                              font=('Segoe UI', 10),
                              relief=tk.FLAT, cursor='hand2',
                              padx=20, pady=10)
        delete_btn.pack(side=tk.LEFT, padx=(0, 10))

        export_btn = tk.Button(button_frame, text="📊 Экспорт",
                              command=self.export_to_excel,
                              bg='#2196F3', fg='white',
                              font=('Segoe UI', 10),
                              relief=tk.FLAT, cursor='hand2',
                              padx=20, pady=10)
        export_btn.pack(side=tk.LEFT)

        # Панель поиска и фильтрации
        filter_panel = tk.Frame(main_container, bg='#ffffff', relief=tk.RAISED, bd=1)
        filter_panel.pack(fill=tk.X, pady=(0, 20), padx=2)

        inner_filter = tk.Frame(filter_panel, bg='#ffffff')
        inner_filter.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(inner_filter, text="🔍 Поиск:",
                bg='#ffffff', font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))

        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_assets())

        search_entry = tk.Entry(inner_filter, textvariable=self.search_var,
                               font=('Segoe UI', 10), width=30)
        search_entry.pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(inner_filter, text="Категория:",
                bg='#ffffff', font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))

        self.category_var = tk.StringVar(value="Все")
        categories = ["Все"] + self.get_categories()
        category_combo = ttk.Combobox(inner_filter, textvariable=self.category_var,
                                     values=categories, width=15, state='readonly')
        category_combo.pack(side=tk.LEFT, padx=(0, 20))
        category_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        tk.Label(inner_filter, text="Статус:",
                bg='#ffffff', font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))

        self.status_var = tk.StringVar(value="Все")
        statuses = ["Все", "Активен", "В ремонте", "Списан"]
        status_combo = ttk.Combobox(inner_filter, textvariable=self.status_var,
                                   values=statuses, width=15, state='readonly')
        status_combo.pack(side=tk.LEFT)
        status_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_assets())

        # Таблица
        table_frame = tk.Frame(main_container, bg='#ffffff', relief=tk.RAISED, bd=1)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, columns=(
            'inventory', 'name', 'category', 'location',
            'responsible', 'cost', 'status', 'date'
        ), show='headings', selectmode='extended')

        columns = {
            'inventory': ('Инв. номер', 100),
            'name': ('Наименование', 200),
            'category': ('Категория', 120),
            'location': ('Расположение', 150),
            'responsible': ('Ответственный', 150),
            'cost': ('Стоимость', 100),
            'status': ('Статус', 100),
            'date': ('Дата покупки', 100)
        }

        for col, (text, width) in columns.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor='center')

        vscroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hscroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)

        self.tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vscroll.grid(row=0, column=1, sticky='ns')
        hscroll.grid(row=1, column=0, sticky='ew')

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind('<Double-Button-1>', lambda e: self.edit_asset())

        # Статистика
        stats_panel = tk.Frame(main_container, bg='#f0f0f0')
        stats_panel.pack(fill=tk.X, pady=(10, 0))

        self.stats_label = tk.Label(stats_panel, text="",
                                   bg='#f0f0f0', font=('Segoe UI', 9))
        self.stats_label.pack(side=tk.LEFT)

        self.update_table()
        self.update_stats()

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

        headers = ['Инв. номер', 'Наименование', 'Категория', 'Расположение',
                  'Ответственный', 'Стоимость', 'Статус', 'Дата покупки', 'Описание']

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF', size=12)
            cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(bottom=Side(style='thin'))

        widths = [15, 25, 15, 20, 20, 12, 12, 15, 30]
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
                        'category': row[2] or '',
                        'location': row[3] or '',
                        'responsible': row[4] or '',
                        'cost': row[5] or 0,
                        'status': row[6] or 'Активен',
                        'date': row[7] or '',
                        'description': row[8] or ''
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

            headers = ['Инв. номер', 'Наименование', 'Категория', 'Расположение',
                      'Ответственный', 'Стоимость', 'Статус', 'Дата покупки', 'Описание']

            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF', size=12)
                cell.fill = PatternFill(start_color='2196F3', end_color='2196F3', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')

            for row_idx, asset in enumerate(self.assets, 2):
                ws.cell(row=row_idx, column=1, value=asset['inventory'])
                ws.cell(row=row_idx, column=2, value=asset['name'])
                ws.cell(row=row_idx, column=3, value=asset['category'])
                ws.cell(row=row_idx, column=4, value=asset['location'])
                ws.cell(row=row_idx, column=5, value=asset['responsible'])
                ws.cell(row=row_idx, column=6, value=asset['cost'])
                ws.cell(row=row_idx, column=7, value=asset['status'])
                ws.cell(row=row_idx, column=8, value=asset['date'])
                ws.cell(row=row_idx, column=9, value=asset['description'])

                if row_idx % 2 == 0:
                    for col in range(1, 10):
                        ws.cell(row=row_idx, column=col).fill = PatternFill(
                            start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')

            widths = [15, 25, 15, 20, 20, 12, 12, 15, 30]
            for col, width in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(col)].width = width

            ws.auto_filter.ref = f"A1:I{len(self.assets) + 1}"
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

        if asset_index is not None:
            dialog = AssetDialog(self.root, "Редактировать актив", self.assets[asset_index])
            self.root.wait_window(dialog.dialog)

            if dialog.result:
                self.assets[asset_index] = dialog.result
                self.save_to_excel()
                self.filter_assets()
                messagebox.showinfo("Успех", "Актив успешно обновлен")

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

    def get_categories(self):
        categories = set()
        for asset in self.assets:
            if asset['category']:
                categories.add(asset['category'])
        return sorted(list(categories))

    def filter_assets(self):
        search_text = self.search_var.get().lower()
        category = self.category_var.get()
        status = self.status_var.get()

        self.filtered_assets = []
        for asset in self.assets:
            if search_text:
                searchable = f"{asset['name']} {asset['inventory']} {asset['location']} {asset['responsible']}"
                if search_text not in searchable.lower():
                    continue

            if category != "Все" and asset['category'] != category:
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
                asset['category'],
                asset['location'],
                asset['responsible'],
                asset['cost'],
                asset['status'],
                asset['date']
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
        self.root.after(300000, self.auto_save)  # 5 минут

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
        self.dialog.geometry("500x600")
        self.dialog.configure(bg='#f0f0f0')
        self.dialog.resizable(False, False)

        self.result = None
        self.asset = asset or {}

        self.setup_ui()

        self.dialog.transient(parent)
        self.dialog.grab_set()

    def setup_ui(self):
        title_label = tk.Label(self.dialog, text="Информация об активе",
                              font=('Segoe UI', 16, 'bold'),
                              bg='#f0f0f0', fg='#333333')
        title_label.pack(pady=20)

        form_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        form_frame.pack(fill=tk.BOTH, expand=True, padx=30)

        fields = [
            ('Инвентарный номер:', 'inventory', False),
            ('Наименование:', 'name', True),
            ('Категория:', 'category', True),
            ('Расположение:', 'location', True),
            ('Ответственный:', 'responsible', True),
            ('Стоимость:', 'cost', False),
            ('Статус:', 'status', False),
            ('Дата покупки:', 'date', False),
            ('Описание:', 'description', False)
        ]

        self.entries = {}

        for i, (label, key, required) in enumerate(fields):
            lbl = tk.Label(form_frame, text=label,
                          bg='#f0f0f0', font=('Segoe UI', 10))
            lbl.grid(row=i, column=0, sticky='w', pady=5)

            if key == 'status':
                var = tk.StringVar(value=self.asset.get(key, 'Активен'))
                entry = ttk.Combobox(form_frame, textvariable=var,
                                    values=['Активен', 'В ремонте', 'Списан'],
                                    state='readonly', width=25)
            elif key == 'date':
                var = tk.StringVar(value=self.asset.get(key, datetime.now().strftime('%d.%m.%Y')))
                entry = tk.Entry(form_frame, textvariable=var, width=27,
                               font=('Segoe UI', 10))
            else:
                var = tk.StringVar(value=self.asset.get(key, ''))
                entry = tk.Entry(form_frame, textvariable=var, width=27,
                               font=('Segoe UI', 10))

            entry.grid(row=i, column=1, sticky='w', pady=5, padx=(10, 0))
            self.entries[key] = var

            if required:
                req_label = tk.Label(form_frame, text="*", fg='red',
                                    bg='#f0f0f0', font=('Segoe UI', 10, 'bold'))
                req_label.grid(row=i, column=2, sticky='w')

        button_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        button_frame.pack(pady=20)

        save_btn = tk.Button(button_frame, text="💾 Сохранить",
                            command=self.save,
                            bg='#4CAF50', fg='white',
                            font=('Segoe UI', 10, 'bold'),
                            relief=tk.FLAT, cursor='hand2',
                            padx=20, pady=10)
        save_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(button_frame, text="Отмена",
                              command=self.dialog.destroy,
                              bg='#9E9E9E', fg='white',
                              font=('Segoe UI', 10),
                              relief=tk.FLAT, cursor='hand2',
                              padx=20, pady=10)
        cancel_btn.pack(side=tk.LEFT, padx=10)

    def save(self):
        if not self.entries['name'].get().strip():
            messagebox.showwarning("Внимание", "Наименование обязательно для заполнения")
            return

        if not self.entries['category'].get().strip():
            messagebox.showwarning("Внимание", "Категория обязательна для заполнения")
            return

        if not self.entries['location'].get().strip():
            messagebox.showwarning("Внимание", "Расположение обязательно для заполнения")
            return

        if not self.entries['responsible'].get().strip():
            messagebox.showwarning("Внимание", "Ответственный обязателен для заполнения")
            return

        self.result = {
            'inventory': self.entries['inventory'].get().strip(),
            'name': self.entries['name'].get().strip(),
            'category': self.entries['category'].get().strip(),
            'location': self.entries['location'].get().strip(),
            'responsible': self.entries['responsible'].get().strip(),
            'cost': self.parse_cost(self.entries['cost'].get()),
            'status': self.entries['status'].get(),
            'date': self.entries['date'].get().strip(),
            'description': self.entries['description'].get().strip()
        }

        self.dialog.destroy()

    def parse_cost(self, value):
        try:
            value = value.replace(' ', '').replace(',', '.')
            return float(value) if value else 0
        except:
            return 0


def main():
    root = tk.Tk()
    app = AssetManager(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()

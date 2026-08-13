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
        """Возвращает словарь цветов в зависимости от текущей темы"""
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
                'button_bg': '#4CAF50',  # базовый для кнопок, но ниже переопределим
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
        """Настройка базовых стилей ttk (будет переопределено в apply_theme)"""
        self.style = ttk.Style()
        self.style.theme_use('default')   # тема с видимой сеткой

    def setup_ui(self):
        """Создание пользовательского интерфейса"""
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

        # Кнопки действий (обычные tk.Button)
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

        # Кнопка переключения темы (правый нижний угол)
        self.theme_btn = tk.Button(self.stats_panel, text="🌙 Тёмная тема",
                                   command=self.toggle_theme, relief=tk.FLAT,
                                   cursor='hand2', padx=10, pady=5)
        self.theme_btn.pack(side=tk.RIGHT)

        self.update_table()
        self.update_stats()

    def apply_theme(self):
        """Применяет текущую тему ко всем элементам"""
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
        """Переключает тему"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    # ----- Все остальные методы остаются без изменений -----
    # (setup_shortcuts, focus_search, create_excel_file, load_assets, save_to_excel,
    #  add_asset, edit_asset, delete_asset, generate_inventory_number,
    #  get_directions, update_filter_values, filter_assets, update_table,
    #  update_stats, export_to_excel, auto_save, on_closing)
    # ... (вставьте их из предыдущего кода без изменений)


class AssetDialog:
    # ----- Класс диалога остаётся прежним -----
    # ... (без изменений)


def main():
    root = tk.Tk()
    app = AssetManager(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()

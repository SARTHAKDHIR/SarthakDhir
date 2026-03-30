import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# ── Colour palette (clean professional light) ────────────────────────────────
ROOT_BG     = '#f0f4f8'   # soft blue-gray canvas
PANEL_BG    = '#ffffff'   # white panels / footer
WIDGET_BG   = '#f8fafc'   # input field background
BORDER      = '#cbd5e1'   # subtle slate border
ACCENT      = '#2563eb'   # professional blue
TEXT_PRI    = '#1e293b'   # near-black primary text
TEXT_SEC    = '#64748b'   # muted secondary text
SUCCESS     = '#16a34a'   # green
ERROR       = '#dc2626'   # red
WARNING     = '#d97706'   # amber
# ─────────────────────────────────────────────────────────────────────────────


def _apply_style():
    style = ttk.Style()
    style.theme_use('clam')

    style.configure('.',
                    background=ROOT_BG,
                    foreground=TEXT_PRI,
                    fieldbackground=WIDGET_BG,
                    troughcolor=WIDGET_BG,
                    bordercolor=BORDER,
                    darkcolor=PANEL_BG,
                    lightcolor=PANEL_BG)

    style.configure('TFrame',      background=ROOT_BG)
    style.configure('Inner.TFrame', background=PANEL_BG)

    style.configure('TLabel',
                    background=ROOT_BG,
                    foreground=TEXT_PRI,
                    font=('Segoe UI', 9))

    style.configure('Title.TLabel',
                    background=ROOT_BG,
                    foreground=ACCENT,
                    font=('Segoe UI', 16, 'bold'))

    style.configure('Footer.TLabel',
                    background=PANEL_BG,
                    foreground=TEXT_SEC,
                    font=('Segoe UI', 8))

    style.configure('Status.TLabel',
                    background=PANEL_BG,
                    foreground=TEXT_SEC,
                    font=('Segoe UI', 9))

    style.configure('TLabelframe',
                    background=ROOT_BG,
                    foreground=ACCENT,
                    bordercolor=BORDER,
                    relief='flat')

    style.configure('TLabelframe.Label',
                    background=ROOT_BG,
                    foreground=ACCENT,
                    font=('Segoe UI', 9, 'bold'))

    style.configure('TButton',
                    background=ACCENT,
                    foreground='#ffffff',
                    font=('Segoe UI', 9, 'bold'),
                    relief='flat',
                    padding=(10, 4))
    style.map('TButton',
              background=[('active', '#1d4ed8'), ('disabled', '#e2e8f0')],
              foreground=[('disabled', TEXT_SEC)])

    style.configure('TEntry',
                    fieldbackground=WIDGET_BG,
                    foreground=TEXT_PRI,
                    insertcolor=TEXT_PRI,
                    bordercolor=BORDER,
                    relief='flat')

    style.configure('TCombobox',
                    fieldbackground=WIDGET_BG,
                    background=WIDGET_BG,
                    foreground=TEXT_PRI,
                    arrowcolor=ACCENT,
                    bordercolor=BORDER,
                    relief='flat',
                    selectbackground=ACCENT,
                    selectforeground='#ffffff')
    style.map('TCombobox',
              fieldbackground=[('readonly', WIDGET_BG)],
              foreground=[('readonly', TEXT_PRI)])

    style.configure('TScrollbar',
                    background=WIDGET_BG,
                    troughcolor=ROOT_BG,
                    arrowcolor=TEXT_SEC,
                    bordercolor=ROOT_BG,
                    relief='flat')

    style.configure('Footer.TFrame', background=PANEL_BG, relief='flat')


class DemandAllocationStudio:
    def __init__(self, root):
        self.root = root
        self.root.title("Demand Allocation Studio  v2.0")
        self.root.geometry("1040x820")
        self.root.configure(bg=ROOT_BG)
        self.root.resizable(True, True)

        self.df = None
        self.file_path = ""
        self.columns = []
        self.available_sheets = []

        self.demand_column_frames = []
        self.demand_column_vars   = []
        self.demand_column_combos = []
        self.demand_month_vars    = []

        _apply_style()
        self.create_menu()
        self.setup_ui()

    # ─────────────────────────────────────────────────────────────────────────
    # MENU
    # ─────────────────────────────────────────────────────────────────────────

    def create_menu(self):
        menubar = tk.Menu(self.root,
                          bg=PANEL_BG, fg=TEXT_PRI,
                          activebackground=ACCENT, activeforeground='#0d1117',
                          relief='flat', bd=0)
        self.root.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0,
                            bg=PANEL_BG, fg=TEXT_PRI,
                            activebackground=ACCENT, activeforeground='#0d1117')
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def show_about(self):
        about_text = """Demand Allocation Studio

Version 2.0

Developed by: Sarthak Dhir

Flexible demand allocation across multiple months
for main and alternate resources, with shortage tracking.

v2.0 highlights:
 • Month-wise allocation — shared resources distributed
   fairly across all StyleDemand classes per month
 • Excel sheet selector — pick any tab from multi-
   sheet workbooks"""
        messagebox.showinfo("About  –  Demand Allocation Studio  v2.0", about_text)

    # ─────────────────────────────────────────────────────────────────────────
    # UI SETUP
    # ─────────────────────────────────────────────────────────────────────────

    def setup_ui(self):
        # Scrollable canvas
        main_canvas = tk.Canvas(self.root, bg=ROOT_BG, highlightthickness=0)
        scrollbar   = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer bar
        footer_frame = ttk.Frame(self.root, style='Footer.TFrame', padding="6")
        footer_frame.pack(side="bottom", fill="x")
        ttk.Label(
            footer_frame,
            text="Demand Allocation Studio  v2.0   |   Developed by Sarthak Dhir",
            style='Footer.TLabel',
            anchor='center'
        ).pack()

        # Main content
        main_frame = ttk.Frame(scrollable_frame, padding="16")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame,
                  text="Demand Allocation Studio  v2.0",
                  style='Title.TLabel').pack(pady=(0, 20))

        self.create_file_section(main_frame)
        self.create_styledemand_section(main_frame)
        self.create_demand_columns_section(main_frame)
        self.create_dial_configuration_section(main_frame)

        self.process_btn = ttk.Button(main_frame, text="Run Allocation",
                                      command=self.process_allocation, state='disabled')
        self.process_btn.pack(pady=20)

        self.create_results_section(main_frame)

        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # ─────────────────────────────────────────────────────────────────────────
    # FILE SECTION
    # ─────────────────────────────────────────────────────────────────────────

    def create_file_section(self, parent):
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding="10")
        file_frame.pack(fill="x", pady=8)

        file_inner = ttk.Frame(file_frame)
        file_inner.pack(fill="x")

        ttk.Label(file_inner, text="Select File:").pack(side="left", padx=(0, 10))
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_inner, textvariable=self.file_path_var, state='readonly').pack(
            side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(file_inner, text="Browse", command=self.browse_file).pack(side="right")

        # Sheet selector (shown only for multi-sheet Excel files)
        self.sheet_frame = ttk.Frame(file_frame)
        self.sheet_frame.pack(fill="x", pady=(6, 0))

        ttk.Label(self.sheet_frame, text="Sheet:").pack(side="left", padx=(0, 10))
        self.sheet_var   = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.sheet_frame, textvariable=self.sheet_var,
                                        state='readonly', width=30)
        self.sheet_combo.pack(side="left")
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        self.sheet_frame.pack_forget()

        self.file_info_label = ttk.Label(file_frame, text="No file selected",
                                         style='Status.TLabel')
        self.file_info_label.pack(anchor="w", pady=(5, 0))

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[('Excel files', '*.xlsx *.xls'),
                       ('CSV files', '*.csv'),
                       ('All files', '*.*')]
        )
        if filename:
            self.file_path = filename
            self.file_path_var.set(filename)
            self._detect_sheets_and_load()

    def _detect_sheets_and_load(self):
        file_ext = os.path.splitext(self.file_path)[1].lower()

        if file_ext in ['.xlsx', '.xls']:
            try:
                xl = pd.ExcelFile(self.file_path)
                self.available_sheets = xl.sheet_names

                if len(self.available_sheets) > 1:
                    self.sheet_combo['values'] = self.available_sheets
                    self.sheet_var.set(self.available_sheets[0])
                    self.sheet_frame.pack(fill="x", pady=(6, 0), before=self.file_info_label)
                    self.file_info_label.config(
                        text=f"Found {len(self.available_sheets)} sheets — please select one.",
                        foreground=ACCENT
                    )
                else:
                    self.sheet_frame.pack_forget()
                    self.sheet_var.set(self.available_sheets[0] if self.available_sheets else '')

                self.load_file()

            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel file: {str(e)}")
        else:
            self.sheet_frame.pack_forget()
            self.available_sheets = []
            self.sheet_var.set('')
            self.load_file()

    def on_sheet_selected(self, event=None):
        self.load_file()

    def load_file(self):
        try:
            file_ext = os.path.splitext(self.file_path)[1].lower()

            if file_ext in ['.xlsx', '.xls']:
                sheet = self.sheet_var.get() or 0
                self.df = pd.read_excel(self.file_path, sheet_name=sheet)
            elif file_ext == '.csv':
                self.df = pd.read_csv(self.file_path)
            else:
                raise ValueError("Unsupported file format")

            rows, cols = self.df.shape
            sheet_info = f"  (Sheet: {self.sheet_var.get()})" if self.sheet_var.get() else ""
            self.file_info_label.config(
                text=f"Loaded: {rows} rows × {cols} columns{sheet_info}",
                foreground=SUCCESS
            )
            self.columns = list(self.df.columns)
            self.populate_column_dropdowns()
            self.process_btn.config(state='normal')

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            self.file_info_label.config(text=f"Error: {str(e)}", foreground=ERROR)

    # ─────────────────────────────────────────────────────────────────────────
    # STYLEDEMAND SECTION
    # ─────────────────────────────────────────────────────────────────────────

    def create_styledemand_section(self, parent):
        styledemand_frame = ttk.LabelFrame(parent, text="StyleDemand Class Configuration", padding="10")
        styledemand_frame.pack(fill="x", pady=8)

        inner_frame = ttk.Frame(styledemand_frame)
        inner_frame.pack(fill="x")

        ttk.Label(inner_frame, text="StyleDemand Class Column:").pack(side="left", padx=(0, 10))
        self.styledemand_var   = tk.StringVar()
        self.styledemand_combo = ttk.Combobox(inner_frame, textvariable=self.styledemand_var,
                                              state='readonly', width=30)
        self.styledemand_combo.pack(side="left", padx=(0, 10))

    # ─────────────────────────────────────────────────────────────────────────
    # DEMAND COLUMNS SECTION
    # ─────────────────────────────────────────────────────────────────────────

    def create_demand_columns_section(self, parent):
        self.demand_frame = ttk.LabelFrame(parent, text="Demand Columns (in priority order)", padding="10")
        self.demand_frame.pack(fill="x", pady=8)

        control_frame = ttk.Frame(self.demand_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(control_frame, text="+ Add Demand Column",
                   command=self.add_demand_column).pack(side="left", padx=(0, 10))
        ttk.Button(control_frame, text="− Remove Last Column",
                   command=self.remove_demand_column).pack(side="left")

        self.demand_columns_container = ttk.Frame(self.demand_frame)
        self.demand_columns_container.pack(fill="x")
        self.add_demand_column()

    def add_demand_column(self):
        column_num = len(self.demand_column_frames) + 1
        col_frame  = ttk.Frame(self.demand_columns_container)
        col_frame.pack(fill="x", pady=2)

        ttk.Label(col_frame, text=f"Month {column_num} Name:").pack(side="left", padx=(0, 10))
        month_var = tk.StringVar()
        ttk.Entry(col_frame, textvariable=month_var, width=15).pack(side="left", padx=(0, 20))

        ttk.Label(col_frame, text="Column:").pack(side="left", padx=(0, 10))
        column_var   = tk.StringVar()
        column_combo = ttk.Combobox(col_frame, textvariable=column_var, state='readonly', width=30)
        column_combo.pack(side="left")

        if self.columns:
            column_combo['values'] = [''] + self.columns

        self.demand_column_frames.append(col_frame)
        self.demand_month_vars.append(month_var)
        self.demand_column_vars.append(column_var)
        self.demand_column_combos.append(column_combo)

        default_months = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug']
        if column_num <= len(default_months):
            month_var.set(default_months[column_num - 1])

    def remove_demand_column(self):
        if len(self.demand_column_frames) > 1:
            self.demand_column_frames.pop().destroy()
            self.demand_month_vars.pop()
            self.demand_column_vars.pop()
            self.demand_column_combos.pop()

    # ─────────────────────────────────────────────────────────────────────────
    # DIAL CONFIGURATION SECTION
    # ─────────────────────────────────────────────────────────────────────────

    def create_dial_configuration_section(self, parent):
        dial_frame = ttk.LabelFrame(parent, text="Resource Configuration", padding="10")
        dial_frame.pack(fill="x", pady=8)

        main_section = ttk.LabelFrame(dial_frame, text="Main Resource", padding="10")
        main_section.pack(fill="x", pady=(0, 10))
        main_inner = ttk.Frame(main_section)
        main_inner.pack(fill="x")

        ttk.Label(main_inner, text="Main Resource Column:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.main_dial_var   = tk.StringVar()
        self.main_dial_combo = ttk.Combobox(main_inner, textvariable=self.main_dial_var,
                                            state='readonly', width=25)
        self.main_dial_combo.grid(row=0, column=1, padx=(0, 20))

        ttk.Label(main_inner, text="Main Availability Column:").grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.main_avl_var   = tk.StringVar()
        self.main_avl_combo = ttk.Combobox(main_inner, textvariable=self.main_avl_var,
                                           state='readonly', width=25)
        self.main_avl_combo.grid(row=0, column=3)

        alt_section = ttk.LabelFrame(dial_frame, text="Alternate Resource", padding="10")
        alt_section.pack(fill="x")
        alt_inner = ttk.Frame(alt_section)
        alt_inner.pack(fill="x")

        ttk.Label(alt_inner, text="Alt Resource Column:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.alt_dial_var   = tk.StringVar()
        self.alt_dial_combo = ttk.Combobox(alt_inner, textvariable=self.alt_dial_var,
                                           state='readonly', width=25)
        self.alt_dial_combo.grid(row=0, column=1, padx=(0, 20))

        ttk.Label(alt_inner, text="Alt Availability Column:").grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.alt_avl_var   = tk.StringVar()
        self.alt_avl_combo = ttk.Combobox(alt_inner, textvariable=self.alt_avl_var,
                                          state='readonly', width=25)
        self.alt_avl_combo.grid(row=0, column=3)

    # ─────────────────────────────────────────────────────────────────────────
    # RESULTS SECTION
    # ─────────────────────────────────────────────────────────────────────────

    def create_results_section(self, parent):
        results_frame = ttk.LabelFrame(parent, text="Results", padding="10")
        results_frame.pack(fill="both", expand=True, pady=8)

        text_frame = ttk.Frame(results_frame)
        text_frame.pack(fill="both", expand=True)

        self.results_text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            height=15,
            bg=WIDGET_BG,
            fg=TEXT_PRI,
            insertbackground=TEXT_PRI,
            selectbackground=ACCENT,
            selectforeground='#ffffff',
            relief='flat',
            bd=1,
            font=('Consolas', 9),
            padx=6,
            pady=6
        )
        results_scrollbar = ttk.Scrollbar(text_frame, orient="vertical",
                                          command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scrollbar.set)
        self.results_text.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")

        self.export_btn = ttk.Button(results_frame, text="Export Results",
                                     command=self.export_results, state='disabled')
        self.export_btn.pack(pady=(10, 0))

    def populate_column_dropdowns(self):
        options = [''] + self.columns
        self.styledemand_combo['values'] = options
        for combo in self.demand_column_combos:
            combo['values'] = options
        self.main_dial_combo['values'] = options
        self.main_avl_combo['values']  = options
        self.alt_dial_combo['values']  = options
        self.alt_avl_combo['values']   = options

    # ─────────────────────────────────────────────────────────────────────────
    # VALIDATION
    # ─────────────────────────────────────────────────────────────────────────

    def validate_inputs(self):
        if self.df is None:
            messagebox.showerror("Error", "Please select and load a file first")
            return False
        if not self.styledemand_var.get():
            messagebox.showerror("Error", "Please select StyleDemand Class Column")
            return False
        demand_cols   = [var.get() for var in self.demand_column_vars  if var.get()]
        demand_months = [var.get() for var in self.demand_month_vars   if var.get()]
        if len(demand_cols) == 0:
            messagebox.showerror("Error", "Please select at least one demand column")
            return False
        if len(demand_months) != len(demand_cols):
            messagebox.showerror("Error", "Please provide month names for all selected demand columns")
            return False
        if not self.main_dial_var.get():
            messagebox.showerror("Error", "Please select Main Resource Column")
            return False
        if not self.main_avl_var.get():
            messagebox.showerror("Error", "Please select Main Availability Column")
            return False
        if not self.alt_dial_var.get():
            messagebox.showerror("Error", "Please select Alt Resource Column")
            return False
        if not self.alt_avl_var.get():
            messagebox.showerror("Error", "Please select Alt Availability Column")
            return False
        return True

    def _is_invalid_avl(self, value):
        if pd.isna(value):
            return True
        return str(value).strip().upper() in ['#N/A', '#NA', 'N/A', 'NA', 'NULL', '', 'NAN']

    # ─────────────────────────────────────────────────────────────────────────
    # CORE PROCESSING
    # ─────────────────────────────────────────────────────────────────────────

    def process_allocation(self):
        if not self.validate_inputs():
            return

        try:
            self.results_text.delete(1.0, tk.END)

            styledemand_col = self.styledemand_var.get()
            demand_cols     = [var.get() for var in self.demand_column_vars  if var.get()]
            demand_months   = [var.get().strip() for var in self.demand_month_vars if var.get()]
            main_dial_col   = self.main_dial_var.get()
            main_avl_col    = self.main_avl_var.get()
            alt_dial_col    = self.alt_dial_var.get()
            alt_avl_col     = self.alt_avl_var.get()

            # ── STEP 1: Build global resource pool (each item counted ONCE) ──
            dial_pool = {}

            for _, row in self.df.iterrows():
                main_dial = row[main_dial_col]
                if pd.notna(main_dial) and main_dial not in dial_pool:
                    avl = float(row[main_avl_col]) if not self._is_invalid_avl(row[main_avl_col]) else 0
                    dial_pool[main_dial] = avl

                alt_dial  = row[alt_dial_col]
                alt_value = row[alt_avl_col]
                if (pd.notna(alt_dial)
                        and alt_dial not in dial_pool
                        and not self._is_invalid_avl(alt_value)):
                    try:
                        dial_pool[alt_dial] = float(alt_value)
                    except (ValueError, TypeError):
                        dial_pool[alt_dial] = 0

            # Log configuration
            self.results_text.insert(tk.END, "=== PROCESSING CONFIGURATION ===\n")
            self.results_text.insert(tk.END, f"StyleDemand Class Column : {styledemand_col}\n")
            self.results_text.insert(tk.END, f"Demand Months (Priority) : {demand_months}\n")
            self.results_text.insert(tk.END, f"Demand Columns           : {demand_cols}\n")
            self.results_text.insert(tk.END, f"Main Resource : {main_dial_col}  |  AVL col: {main_avl_col}\n")
            self.results_text.insert(tk.END, f"Alt  Resource : {alt_dial_col}   |  AVL col: {alt_avl_col}\n")

            self.results_text.insert(tk.END, f"\n--- Resource Pool ({len(dial_pool)} unique items) ---\n")
            for dial, avl in dial_pool.items():
                self.results_text.insert(tk.END, f"  {dial}  →  AVL = {avl}\n")

            # ── STEP 2: Pre-build results store ─────────────────────────────
            allocation_store = {
                idx: {month: {'main': 0, 'alt': 0, 'shortage': 0} for month in demand_months}
                for idx in self.df.index
            }

            # ── STEP 3: MONTH-WISE allocation loop ───────────────────────────
            #   Outer: each MONTH  →  Inner: each ROW
            #   All SDCs compete fairly per month before next month starts
            self.results_text.insert(tk.END, "\n=== PROCESSING RESULTS (Month-wise) ===\n")

            for month_idx, (month, demand_col) in enumerate(zip(demand_months, demand_cols)):
                self.results_text.insert(tk.END, f"\n--- Month: {month} ---\n")

                for idx, row in self.df.iterrows():
                    demand = float(row[demand_col]) if not self._is_invalid_avl(row[demand_col]) else 0
                    if demand == 0:
                        continue

                    main_dial = row[main_dial_col]
                    alt_dial  = row[alt_dial_col]

                    remaining_main = dial_pool.get(main_dial, 0)

                    if self._is_invalid_avl(row[alt_avl_col]):
                        remaining_alt = 0
                        alt_has_dial  = False
                    else:
                        remaining_alt = dial_pool.get(alt_dial, 0)
                        alt_has_dial  = True

                    if remaining_main <= remaining_alt:
                        main_alloc = min(demand, remaining_main)
                        alt_alloc  = min(demand - main_alloc, remaining_alt)
                    else:
                        alt_alloc  = min(demand, remaining_alt)
                        main_alloc = min(demand - alt_alloc, remaining_main)

                    shortage = max(0, demand - main_alloc - alt_alloc)

                    allocation_store[idx][month]['main']     = main_alloc
                    allocation_store[idx][month]['alt']      = alt_alloc
                    allocation_store[idx][month]['shortage'] = shortage

                    dial_pool[main_dial] = remaining_main - main_alloc
                    if alt_has_dial and alt_dial in dial_pool:
                        dial_pool[alt_dial] = remaining_alt - alt_alloc

                    if month_idx < 2 and idx < 3:
                        alt_status = (
                            "No Alt Resource" if not alt_has_dial
                            else f"Alt({alt_dial}) left={dial_pool.get(alt_dial, 0):.0f}"
                        )
                        self.results_text.insert(tk.END,
                            f"  Row {idx+1} [{row[styledemand_col]}] "
                            f"Demand={demand}  "
                            f"Main({main_dial})={main_alloc} left={dial_pool[main_dial]:.0f}  "
                            f"{alt_status}  Alt={alt_alloc}  Shortage={shortage}\n"
                        )

            # ── STEP 4: Assemble final DataFrame ─────────────────────────────
            results = []
            for idx, row in self.df.iterrows():
                result_row = row.copy()
                for month in demand_months:
                    result_row[f'Main {month} Allocation'] = allocation_store[idx][month]['main']
                    result_row[f'Alt {month} Allocation']  = allocation_store[idx][month]['alt']
                    result_row[f'{month} Shortage']        = allocation_store[idx][month]['shortage']
                results.append(result_row)

            self.results_df = pd.DataFrame(results)

            # Summary
            total_shortage = 0
            shortage_rows  = 0
            for _, result_row in self.results_df.iterrows():
                row_shortage = sum(result_row.get(f'{m} Shortage', 0) for m in demand_months)
                if row_shortage > 0:
                    shortage_rows  += 1
                    total_shortage += row_shortage

            self.results_text.insert(tk.END,
                f"\n=== SUMMARY ===\n"
                f"Total rows processed : {len(results)}\n"
                f"Rows with shortage   : {shortage_rows}\n"
                f"Total shortage units : {total_shortage}\n"
                f"\nResource Pool — Remaining after full allocation:\n"
            )
            for dial, avl in dial_pool.items():
                self.results_text.insert(tk.END, f"  {dial}  →  {avl:.0f} remaining\n")

            self.results_text.insert(tk.END, "\nResults ready for export.\n")
            self.export_btn.config(state='normal')

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process allocation: {str(e)}")
            self.results_text.insert(tk.END, f"Error: {str(e)}\n")

    # ─────────────────────────────────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────────────────────────────────

    def export_results(self):
        if not hasattr(self, 'results_df'):
            messagebox.showerror("Error", "No results to export")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save results as",
            defaultextension=".xlsx",
            filetypes=[('Excel files', '*.xlsx'),
                       ('CSV files', '*.csv'),
                       ('All files', '*.*')]
        )

        if not file_path:
            return

        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext == '.csv':
                self.results_df.to_csv(file_path, index=False)
            else:
                self.results_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Results exported to:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results: {str(e)}")


def main():
    root = tk.Tk()
    app  = DemandAllocationStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()

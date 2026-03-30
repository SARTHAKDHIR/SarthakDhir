import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import os
import sys

# Fix encoding for Windows terminal
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


class MasterExcelToolkit:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Demand Toolkit - Master Excel Processing Toolkit")
        self.root.geometry("1000x800")

        # Create main frame
        main_frame = ttk.Frame(root)
        main_frame.pack(fill='both', expand=True)

        # Help Tab Button at the top
        help_button_frame = ttk.Frame(main_frame)
        help_button_frame.pack(fill='x', padx=10, pady=5)

        help_button = ttk.Button(help_button_frame,
                                text="📖 Help & Information",
                                command=self.show_help_window)
        help_button.pack(side='left')

        # Create main notebook for the tools
        self.main_notebook = ttk.Notebook(main_frame)
        self.main_notebook.pack(fill='both', expand=True, padx=10, pady=(5, 10))

        # Tool 1: Vendor Merger
        self.tool1_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.tool1_frame, text="Tool 1: Vendor Merger")
        self.vendor_merger = VendorMergerGUI(self.tool1_frame)

        # Tool 2: Data Combiner & VLOOKUP
        self.tool2_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.tool2_frame, text="Tool 2: Data Combiner & VLOOKUP")
        self.data_combiner = ExcelCombinerApp(self.tool2_frame)

        # Tool 3: Shortage Allocation
        self.tool3_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.tool3_frame, text="Tool 3: Shortage Allocation")
        self.shortage_allocation = ShortageAllocationGUI(self.tool3_frame)

        # Enable mouse wheel scrolling for all canvases
        self.bind_mousewheel(root)

        # Credits footer (always visible at bottom)
        credits_frame = ttk.Frame(root)
        credits_frame.pack(side='bottom', fill='x', padx=5, pady=2)

        credits_label = ttk.Label(credits_frame,
                                 text="Excel Demand Toolkit v1.0  |  Developed by: Sarthak Dhir",
                                 font=('Arial', 8),
                                 foreground='gray')
        credits_label.pack(side='left')

    def bind_mousewheel(self, widget):
        """Enable mouse wheel scrolling for all canvas widgets"""
        def _on_mousewheel(event):
            # Find the canvas widget under the mouse
            widget = event.widget
            while widget:
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1*(event.delta/120)), "units")
                    break
                widget = widget.master

        # Bind for Windows/Mac
        self.root.bind_all("<MouseWheel>", _on_mousewheel)
        # Bind for Linux
        self.root.bind_all("<Button-4>", lambda e: _on_mousewheel(type('Event', (), {'delta': 120, 'widget': e.widget})()))
        self.root.bind_all("<Button-5>", lambda e: _on_mousewheel(type('Event', (), {'delta': -120, 'widget': e.widget})()))

    def show_help_window(self):
        """Show help information in a separate window"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Help & Information - Excel Demand Toolkit")
        help_window.geometry("700x600")

        # Create canvas and scrollbar for help content
        canvas = tk.Canvas(help_window)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel for help window
        def _on_mousewheel_help(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel_help)

        # Main content frame
        content_frame = ttk.Frame(scrollable_frame, padding=40)
        content_frame.pack(fill='both', expand=True)

        # Header
        title = ttk.Label(content_frame,
                         text="Excel Demand Toolkit",
                         font=('Arial', 20, 'bold'))
        title.pack(pady=(0, 10))

        version_info = ttk.Label(content_frame,
                                text="Version 1.0\n\nDeveloped by: Sarthak Dhir",
                                font=('Arial', 11),
                                justify='center')
        version_info.pack(pady=(0, 30))

        ttk.Separator(content_frame, orient='horizontal').pack(fill='x', pady=20)

        # Description
        desc_title = ttk.Label(content_frame,
                              text="About This Tool",
                              font=('Arial', 14, 'bold'))
        desc_title.pack(pady=(0, 15))

        description = """This comprehensive toolkit combines three powerful Excel processing tools designed for Demand Allocation across Vendors:

1) Vendor Merger

2) Data Combiner & VLOOKUP

3) Shortage Allocation

"""

        desc_label = ttk.Label(content_frame,
                              text=description,
                              font=('Arial', 11),
                              justify='left',
                              wraplength=600)
        desc_label.pack(pady=(0, 30))

        ttk.Separator(content_frame, orient='horizontal').pack(fill='x', pady=20)

        # Contact
        contact = ttk.Label(content_frame,
                           text="Excel Demand Toolkit — Master Excel Processing for Demand Allocation",
                           font=('Arial', 10, 'italic'),
                           foreground='blue')
        contact.pack(pady=10)

        # Close button
        close_button = ttk.Button(content_frame,
                                 text="Close",
                                 command=help_window.destroy)
        close_button.pack(pady=20)


# ==================== TOOL 1: VENDOR MERGER ====================
class VendorMergerGUI:
    def __init__(self, parent):
        self.root = parent

        # Variables to store file paths and selections
        self.shortage_file = tk.StringVar()
        self.po_file = tk.StringVar()
        self.pr_file = tk.StringVar()

        self.shortage_sheet = tk.StringVar()
        self.po_sheet = tk.StringVar()
        self.pr_sheet = tk.StringVar()

        self.shortage_key_col = tk.StringVar()
        self.po_item_col = tk.StringVar()
        self.po_vendor_col = tk.StringVar()
        self.pr_item_col = tk.StringVar()
        self.pr_vendor_col = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Create canvas and scrollbar for Tool 1
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Main frame with padding
        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for responsiveness
        scrollable_frame.columnconfigure(0, weight=1)
        scrollable_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0

        # Title
        title_label = ttk.Label(main_frame, text="Excel Vendor Merger Tool",
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=row, column=0, columnspan=3, pady=10)
        row += 1

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0,
                                                           columnspan=3, sticky='ew', pady=5)
        row += 1

        # --- SHORTAGE FILE SECTION ---
        shortage_frame = ttk.LabelFrame(main_frame, text="Plan Shortage File", padding="5")
        shortage_frame.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)

        ttk.Label(shortage_frame, text="File:").grid(row=0, column=0, sticky='w')
        ttk.Entry(shortage_frame, textvariable=self.shortage_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(shortage_frame, text="Browse",
                  command=lambda: self.browse_file(self.shortage_file, 'shortage')).grid(row=0, column=2)

        ttk.Label(shortage_frame, text="Sheet:").grid(row=1, column=0, sticky='w', pady=5)
        self.shortage_sheet_combo = ttk.Combobox(shortage_frame, textvariable=self.shortage_sheet, width=47)
        self.shortage_sheet_combo.grid(row=1, column=1, pady=5)
        self.shortage_sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.load_columns('shortage'))

        ttk.Label(shortage_frame, text="Key Column (Item Code):").grid(row=2, column=0, sticky='w', pady=5)
        self.shortage_col_combo = ttk.Combobox(shortage_frame, textvariable=self.shortage_key_col, width=47)
        self.shortage_col_combo.grid(row=2, column=1, pady=5)

        row += 1

        # --- PO FILE SECTION ---
        po_frame = ttk.LabelFrame(main_frame, text="PO File", padding="5")
        po_frame.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)

        ttk.Label(po_frame, text="File:").grid(row=0, column=0, sticky='w')
        ttk.Entry(po_frame, textvariable=self.po_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(po_frame, text="Browse",
                  command=lambda: self.browse_file(self.po_file, 'po')).grid(row=0, column=2)

        ttk.Label(po_frame, text="Sheet:").grid(row=1, column=0, sticky='w', pady=5)
        self.po_sheet_combo = ttk.Combobox(po_frame, textvariable=self.po_sheet, width=47)
        self.po_sheet_combo.grid(row=1, column=1, pady=5)
        self.po_sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.load_columns('po'))

        ttk.Label(po_frame, text="Item Code Column:").grid(row=2, column=0, sticky='w', pady=5)
        self.po_item_combo = ttk.Combobox(po_frame, textvariable=self.po_item_col, width=47)
        self.po_item_combo.grid(row=2, column=1, pady=5)

        ttk.Label(po_frame, text="Vendor Name Column:").grid(row=3, column=0, sticky='w', pady=5)
        self.po_vendor_combo = ttk.Combobox(po_frame, textvariable=self.po_vendor_col, width=47)
        self.po_vendor_combo.grid(row=3, column=1, pady=5)

        row += 1

        # --- PR FILE SECTION ---
        pr_frame = ttk.LabelFrame(main_frame, text="PR File", padding="5")
        pr_frame.grid(row=row, column=0, columnspan=3, sticky='ew', pady=5)

        ttk.Label(pr_frame, text="File:").grid(row=0, column=0, sticky='w')
        ttk.Entry(pr_frame, textvariable=self.pr_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(pr_frame, text="Browse",
                  command=lambda: self.browse_file(self.pr_file, 'pr')).grid(row=0, column=2)

        ttk.Label(pr_frame, text="Sheet:").grid(row=1, column=0, sticky='w', pady=5)
        self.pr_sheet_combo = ttk.Combobox(pr_frame, textvariable=self.pr_sheet, width=47)
        self.pr_sheet_combo.grid(row=1, column=1, pady=5)
        self.pr_sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.load_columns('pr'))

        ttk.Label(pr_frame, text="Item Code Column:").grid(row=2, column=0, sticky='w', pady=5)
        self.pr_item_combo = ttk.Combobox(pr_frame, textvariable=self.pr_item_col, width=47)
        self.pr_item_combo.grid(row=2, column=1, pady=5)

        ttk.Label(pr_frame, text="Vendor Name Column:").grid(row=3, column=0, sticky='w', pady=5)
        self.pr_vendor_combo = ttk.Combobox(pr_frame, textvariable=self.pr_vendor_col, width=47)
        self.pr_vendor_combo.grid(row=3, column=1, pady=5)

        row += 1

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=row, column=0,
                                                           columnspan=3, sticky='ew', pady=10)
        row += 1

        # Process button
        process_btn = ttk.Button(main_frame, text="Process Files",
                                command=self.process_files, style='Accent.TButton')
        process_btn.grid(row=row, column=0, columnspan=3, pady=10)

        row += 1

        # Status text area
        ttk.Label(main_frame, text="Status:").grid(row=row, column=0, sticky='w')
        row += 1

        # Create text widget with scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=row, column=0, columnspan=3, sticky='ew')

        self.status_text = tk.Text(text_frame, height=8, width=80)
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)

        self.status_text.grid(row=0, column=0, sticky='ew')
        scrollbar.grid(row=0, column=1, sticky='ns')

        text_frame.columnconfigure(0, weight=1)

    def browse_file(self, var, file_type):
        filename = filedialog.askopenfilename(
            title=f"Select {file_type.upper()} Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
            self.load_sheets(file_type)

    def load_sheets(self, file_type):
        try:
            if file_type == 'shortage':
                file_path = self.shortage_file.get()
                combo = self.shortage_sheet_combo
            elif file_type == 'po':
                file_path = self.po_file.get()
                combo = self.po_sheet_combo
            else:  # pr
                file_path = self.pr_file.get()
                combo = self.pr_sheet_combo

            if file_path:
                xls = pd.ExcelFile(file_path)
                combo['values'] = xls.sheet_names
                if xls.sheet_names:
                    combo.current(0)
                    self.load_columns(file_type)
        except Exception as e:
            self.log_status(f"Error loading sheets from {file_type}: {str(e)}")

    def load_columns(self, file_type):
        try:
            if file_type == 'shortage':
                file_path = self.shortage_file.get()
                sheet = self.shortage_sheet.get()
                combo = self.shortage_col_combo
            elif file_type == 'po':
                file_path = self.po_file.get()
                sheet = self.po_sheet.get()
                combos = [self.po_item_combo, self.po_vendor_combo]
            else:  # pr
                file_path = self.pr_file.get()
                sheet = self.pr_sheet.get()
                combos = [self.pr_item_combo, self.pr_vendor_combo]

            if file_path and sheet:
                df = pd.read_excel(file_path, sheet_name=sheet, nrows=0)
                columns = df.columns.tolist()

                if file_type == 'shortage':
                    combo['values'] = columns
                else:
                    for c in combos:
                        c['values'] = columns
        except Exception as e:
            self.log_status(f"Error loading columns from {file_type}: {str(e)}")

    def log_status(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def validate_inputs(self):
        # Check if all files are selected
        if not all([self.shortage_file.get(), self.po_file.get(), self.pr_file.get()]):
            messagebox.showerror("Error", "Please select all three Excel files")
            return False

        # Check if all sheets are selected
        if not all([self.shortage_sheet.get(), self.po_sheet.get(), self.pr_sheet.get()]):
            messagebox.showerror("Error", "Please select sheets for all files")
            return False

        # Check if all columns are selected
        if not all([self.shortage_key_col.get(), self.po_item_col.get(),
                   self.po_vendor_col.get(), self.pr_item_col.get(), self.pr_vendor_col.get()]):
            messagebox.showerror("Error", "Please select all required columns")
            return False

        return True

    def process_files(self):
        if not self.validate_inputs():
            return

        try:
            # Clear status
            self.status_text.delete(1.0, tk.END)
            self.log_status("Starting processing...")

            # --- Load plan shortage ---
            self.log_status(f"Loading shortage file: {os.path.basename(self.shortage_file.get())}")
            df_shortage = pd.read_excel(self.shortage_file.get(), sheet_name=self.shortage_sheet.get())

            # Rename the key column to "Item Code" for consistency
            item_code_col = self.shortage_key_col.get()
            df_shortage = df_shortage.rename(columns={item_code_col: "Item Code"})

            # --- Load PO & PR with only Item Code + Vendor Name ---
            self.log_status(f"Loading PO file: {os.path.basename(self.po_file.get())}")
            df_po = pd.read_excel(self.po_file.get(), sheet_name=self.po_sheet.get(),
                                 usecols=[self.po_item_col.get(), self.po_vendor_col.get()])
            df_po = df_po.rename(columns={
                self.po_item_col.get(): "Item Code",
                self.po_vendor_col.get(): "Vendor Name"
            })

            self.log_status(f"Loading PR file: {os.path.basename(self.pr_file.get())}")
            df_pr = pd.read_excel(self.pr_file.get(), sheet_name=self.pr_sheet.get(),
                                 usecols=[self.pr_item_col.get(), self.pr_vendor_col.get()])
            df_pr = df_pr.rename(columns={
                self.pr_item_col.get(): "Item Code",
                self.pr_vendor_col.get(): "Vendor Name"
            })

            # --- Combine vendors ---
            self.log_status("Combining vendor data...")
            df_vendors = pd.concat([df_po, df_pr]).drop_duplicates()

            # Count vendors per item
            vendor_counts = df_vendors.groupby("Item Code")["Vendor Name"].nunique().reset_index()
            vendor_counts = vendor_counts.rename(columns={"Vendor Name": "Vendor Count"})
            vendor_counts["Supplier Type"] = vendor_counts["Vendor Count"].apply(
                lambda x: "Single" if x == 1 else "Multiple"
            )

            # Final vendor list
            df_vendor_list = pd.merge(df_vendors, vendor_counts, on="Item Code", how="left")

            # --- Merge shortage with vendor info (LEFT JOIN so all items remain) ---
            self.log_status("Merging shortage with vendor information...")
            df_expanded = pd.merge(df_shortage, df_vendor_list, on="Item Code", how="left")

            # Fill missing vendor info
            df_expanded["Vendor Name"] = df_expanded["Vendor Name"].fillna("#N/A")
            df_expanded["Supplier Type"] = df_expanded["Supplier Type"].fillna("None")
            df_expanded["Vendor Count"] = df_expanded["Vendor Count"].fillna(0).astype(int)

            # --- Ensure only one Vendor Name column and Vendor Count ---
            keep_cols = ["Item Code", "Vendor Name", "Vendor Count", "Supplier Type"]
            remaining_cols = [c for c in df_expanded.columns if c not in keep_cols]
            df_final = df_expanded[keep_cols + remaining_cols]

            # Rename Item Code back to original if needed
            if item_code_col != "Item Code":
                df_final = df_final.rename(columns={"Item Code": item_code_col})
                df_vendor_list = df_vendor_list.rename(columns={"Item Code": item_code_col})

            # --- Save outputs ---
            output_shortage = filedialog.asksaveasfilename(
                title="Save shortage with vendors as",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile="shortage_with_vendors.xlsx"
            )

            if output_shortage:
                df_final.to_excel(output_shortage, index=False)
                self.log_status(f"Shortage with vendors saved as: {output_shortage}")

                output_vendors = filedialog.asksaveasfilename(
                    title="Save vendor list as",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    initialfile="vendor_list.xlsx"
                )

                if output_vendors:
                    df_vendor_list.to_excel(output_vendors, index=False)
                    self.log_status(f"Vendor list saved as: {output_vendors}")

                self.log_status("\n✅ Processing completed successfully!")
                messagebox.showinfo("Success", "Files processed and saved successfully!")
            else:
                self.log_status("Save operation cancelled.")

        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.log_status(f"\n❌ {error_msg}")
            messagebox.showerror("Error", error_msg)


# ==================== TOOL 2: DATA COMBINER & VLOOKUP ====================
class ExcelCombinerApp:
    def __init__(self, parent):
        self.root = parent

        # Data storage
        self.main_df = None
        self.po_df = None
        self.pr_df = None
        self.main_file_path = None
        self.po_file_path = None
        self.pr_file_path = None

        # Create canvas and scrollbar for the entire Tool 2
        self.main_canvas = tk.Canvas(parent)
        self.main_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.main_canvas.yview)
        self.main_scrollable_frame = ttk.Frame(self.main_canvas)

        self.main_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )

        self.main_canvas.create_window((0, 0), window=self.main_scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_scrollbar.pack(side="right", fill="y")

        # Create notebook for tabs inside the main scrollable frame
        self.notebook = ttk.Notebook(self.main_scrollable_frame)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Step 1: Combine Columns
        self.step1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.step1_frame, text="Step 1: Combine Columns")
        self.create_step1_ui()

        # Step 2: VLOOKUP from PO/PR
        self.step2_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.step2_frame, text="Step 2: VLOOKUP PO/PR")
        self.create_step2_ui()

        # Step 3: Add PO/PR Columns
        self.step3_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.step3_frame, text="Step 3: Add PO/PR Columns")
        self.create_step3_ui()

        # Step 4: Export
        self.step4_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.step4_frame, text="Step 4: Export")
        self.create_step4_ui()

    def create_step1_ui(self):
        """Step 1: Load file and combine columns"""
        # File selection
        file_frame = ttk.LabelFrame(self.step1_frame, text="Select Main File", padding=10)
        file_frame.pack(fill='x', padx=10, pady=10)

        self.main_file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.main_file_label.pack(side='left', fill='x', expand=True)

        ttk.Button(file_frame, text="Browse", command=self.load_main_file).pack(side='right')

        # Sheet selection (for Excel files with multiple sheets)
        sheet_frame = ttk.LabelFrame(self.step1_frame, text="Select Sheet", padding=10)
        sheet_frame.pack(fill='x', padx=10, pady=10)

        self.main_sheet_combo = ttk.Combobox(sheet_frame, state='disabled')
        self.main_sheet_combo.pack(fill='x')
        ttk.Button(sheet_frame, text="Load Sheet", command=self.load_main_sheet).pack(pady=5)

        # Column combination
        combine_frame = ttk.LabelFrame(self.step1_frame, text="Combine Columns", padding=10)
        combine_frame.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(combine_frame, text="Select First Column:").grid(row=0, column=0, sticky='w', pady=5)
        self.col1_combo = ttk.Combobox(combine_frame, state='disabled')
        self.col1_combo.grid(row=0, column=1, sticky='ew', pady=5, padx=5)

        ttk.Label(combine_frame, text="Select Second Column:").grid(row=1, column=0, sticky='w', pady=5)
        self.col2_combo = ttk.Combobox(combine_frame, state='disabled')
        self.col2_combo.grid(row=1, column=1, sticky='ew', pady=5, padx=5)

        ttk.Label(combine_frame, text="New Column Name:").grid(row=2, column=0, sticky='w', pady=5)
        self.new_col_entry = ttk.Entry(combine_frame)
        self.new_col_entry.insert(0, "Item Vendor")
        self.new_col_entry.grid(row=2, column=1, sticky='ew', pady=5, padx=5)

        ttk.Label(combine_frame, text="Separator (optional):").grid(row=3, column=0, sticky='w', pady=5)
        self.separator_entry = ttk.Entry(combine_frame)
        self.separator_entry.insert(0, "")
        self.separator_entry.grid(row=3, column=1, sticky='ew', pady=5, padx=5)

        combine_frame.columnconfigure(1, weight=1)

        ttk.Button(combine_frame, text="Combine Columns",
                  command=self.combine_columns).grid(row=4, column=0, columnspan=2, pady=10)

        # Status
        self.step1_status = ttk.Label(self.step1_frame, text="", foreground="blue")
        self.step1_status.pack(pady=5)

    def create_step2_ui(self):
        """Step 2: Load PO/PR files and perform VLOOKUP"""
        # PO File
        po_frame = ttk.LabelFrame(self.step2_frame, text="PO File", padding=10)
        po_frame.pack(fill='x', padx=10, pady=10)

        self.po_file_label = ttk.Label(po_frame, text="No file selected", foreground="gray")
        self.po_file_label.pack(side='left', fill='x', expand=True)
        ttk.Button(po_frame, text="Browse", command=self.load_po_file).pack(side='right')

        po_sheet_frame = ttk.Frame(po_frame)
        po_sheet_frame.pack(fill='x', pady=5)
        ttk.Label(po_sheet_frame, text="Sheet:").pack(side='left')
        self.po_sheet_combo = ttk.Combobox(po_sheet_frame, state='disabled')
        self.po_sheet_combo.pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(po_sheet_frame, text="Load", command=self.load_po_sheet).pack(side='right')

        # PR File
        pr_frame = ttk.LabelFrame(self.step2_frame, text="PR File", padding=10)
        pr_frame.pack(fill='x', padx=10, pady=10)

        self.pr_file_label = ttk.Label(pr_frame, text="No file selected", foreground="gray")
        self.pr_file_label.pack(side='left', fill='x', expand=True)
        ttk.Button(pr_frame, text="Browse", command=self.load_pr_file).pack(side='right')

        pr_sheet_frame = ttk.Frame(pr_frame)
        pr_sheet_frame.pack(fill='x', pady=5)
        ttk.Label(pr_sheet_frame, text="Sheet:").pack(side='left')
        self.pr_sheet_combo = ttk.Combobox(pr_sheet_frame, state='disabled')
        self.pr_sheet_combo.pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(pr_sheet_frame, text="Load", command=self.load_pr_sheet).pack(side='right')

        # Column selection for VLOOKUP
        vlookup_frame = ttk.LabelFrame(self.step2_frame, text="VLOOKUP Configuration", padding=10)
        vlookup_frame.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(vlookup_frame, text="Main File Lookup Column:").grid(row=0, column=0, sticky='w', pady=5)
        self.main_lookup_combo = ttk.Combobox(vlookup_frame, state='disabled')
        self.main_lookup_combo.grid(row=0, column=1, sticky='ew', pady=5, padx=5)

        ttk.Label(vlookup_frame, text="PO Lookup Column:").grid(row=1, column=0, sticky='w', pady=5)
        self.po_lookup_combo = ttk.Combobox(vlookup_frame, state='disabled')
        self.po_lookup_combo.grid(row=1, column=1, sticky='ew', pady=5, padx=5)

        ttk.Label(vlookup_frame, text="PR Lookup Column:").grid(row=2, column=0, sticky='w', pady=5)
        self.pr_lookup_combo = ttk.Combobox(vlookup_frame, state='disabled')
        self.pr_lookup_combo.grid(row=2, column=1, sticky='ew', pady=5, padx=5)

        vlookup_frame.columnconfigure(1, weight=1)

        # PO Columns to add
        ttk.Label(vlookup_frame, text="Select PO Columns to Add:").grid(row=3, column=0, sticky='nw', pady=10)
        self.po_columns_frame = ttk.Frame(vlookup_frame)
        self.po_columns_frame.grid(row=3, column=1, sticky='ew', pady=10)
        self.po_column_vars = {}

        # PR Columns to add
        ttk.Label(vlookup_frame, text="Select PR Columns to Add:").grid(row=4, column=0, sticky='nw', pady=10)
        self.pr_columns_frame = ttk.Frame(vlookup_frame)
        self.pr_columns_frame.grid(row=4, column=1, sticky='ew', pady=10)
        self.pr_column_vars = {}

        ttk.Button(vlookup_frame, text="Perform VLOOKUP",
                  command=self.perform_vlookup).grid(row=5, column=0, columnspan=2, pady=10)

        self.step2_status = ttk.Label(self.step2_frame, text="", foreground="blue")
        self.step2_status.pack(pady=5)

    def create_step3_ui(self):
        """Step 3: Add combined PO/PR columns"""
        instructions = ttk.Label(self.step3_frame,
                                text="Add combined PO/PR columns by selecting corresponding months",
                                wraplength=800)
        instructions.pack(pady=10)

        # Create a frame to hold buttons
        button_frame = ttk.Frame(self.step3_frame)
        button_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(button_frame, text="Add Month Combination",
                  command=self.add_month_combination).pack(side='left', padx=5)

        ttk.Button(button_frame, text="Create Combined Columns",
                  command=self.create_combined_columns).pack(side='left', padx=5)

        # Scrollable frame for month combinations
        canvas = tk.Canvas(self.step3_frame)
        scrollbar = ttk.Scrollbar(self.step3_frame, orient="vertical", command=canvas.yview)
        self.month_frame = ttk.Frame(canvas)

        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True, padx=10)
        canvas.create_window((0, 0), window=self.month_frame, anchor="nw")

        self.month_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.month_combinations = []

        self.step3_status = ttk.Label(self.step3_frame, text="", foreground="blue")
        self.step3_status.pack(pady=5)

    def create_step4_ui(self):
        """Step 4: Export final file"""
        export_frame = ttk.LabelFrame(self.step4_frame, text="Export Options", padding=20)
        export_frame.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(export_frame, text="Export Format:").pack(pady=5)
        self.export_format = tk.StringVar(value="xlsx")
        ttk.Radiobutton(export_frame, text="Excel (.xlsx)",
                       variable=self.export_format, value="xlsx").pack()
        ttk.Radiobutton(export_frame, text="CSV (.csv)",
                       variable=self.export_format, value="csv").pack()

        ttk.Button(export_frame, text="Export Final File",
                  command=self.export_file, style='Accent.TButton').pack(pady=20)

        # Preview frame
        preview_frame = ttk.LabelFrame(self.step4_frame, text="Data Preview", padding=10)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.preview_text = tk.Text(preview_frame, height=15, wrap='none')
        preview_scrollbar_y = ttk.Scrollbar(preview_frame, orient='vertical',
                                           command=self.preview_text.yview)
        preview_scrollbar_x = ttk.Scrollbar(preview_frame, orient='horizontal',
                                           command=self.preview_text.xview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar_y.set,
                                   xscrollcommand=preview_scrollbar_x.set)

        preview_scrollbar_y.pack(side='right', fill='y')
        preview_scrollbar_x.pack(side='bottom', fill='x')
        self.preview_text.pack(fill='both', expand=True)

        ttk.Button(preview_frame, text="Refresh Preview",
                  command=self.update_preview).pack(pady=5)

        self.step4_status = ttk.Label(self.step4_frame, text="", foreground="blue")
        self.step4_status.pack(pady=5)

    # Step 1 Methods
    def load_main_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Main File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.main_file_path = file_path
            self.main_file_label.config(text=Path(file_path).name, foreground="black")

            # Check if Excel file with multiple sheets
            if file_path.endswith(('.xlsx', '.xls')):
                try:
                    xl_file = pd.ExcelFile(file_path)
                    sheets = xl_file.sheet_names
                    self.main_sheet_combo['values'] = sheets
                    self.main_sheet_combo.set(sheets[0])
                    self.main_sheet_combo['state'] = 'readonly'
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")
            else:
                # CSV file - load directly
                self.load_main_sheet()

    def load_main_sheet(self):
        if not self.main_file_path:
            messagebox.showwarning("Warning", "Please select a file first")
            return

        try:
            if self.main_file_path.endswith('.csv'):
                self.main_df = pd.read_csv(self.main_file_path)
            else:
                sheet_name = self.main_sheet_combo.get()
                self.main_df = pd.read_excel(self.main_file_path, sheet_name=sheet_name)

            # Populate column dropdowns
            columns = list(self.main_df.columns)
            self.col1_combo['values'] = columns
            self.col2_combo['values'] = columns
            self.col1_combo['state'] = 'readonly'
            self.col2_combo['state'] = 'readonly'

            self.step1_status.config(text=f"Loaded {len(self.main_df)} rows, {len(columns)} columns")
            messagebox.showinfo("Success", f"File loaded successfully!\nRows: {len(self.main_df)}\nColumns: {len(columns)}")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

    def combine_columns(self):
        if self.main_df is None:
            messagebox.showwarning("Warning", "Please load a file first")
            return

        col1 = self.col1_combo.get()
        col2 = self.col2_combo.get()
        new_col_name = self.new_col_entry.get()
        separator = self.separator_entry.get()

        if not col1 or not col2:
            messagebox.showwarning("Warning", "Please select both columns")
            return

        if not new_col_name:
            messagebox.showwarning("Warning", "Please enter a new column name")
            return

        try:
            # Combine columns (convert to string and handle NaN)
            self.main_df[new_col_name] = (
                self.main_df[col1].astype(str).fillna('') +
                separator +
                self.main_df[col2].astype(str).fillna('')
            )

            self.step1_status.config(text=f"Column '{new_col_name}' created successfully!")
            messagebox.showinfo("Success", f"Column '{new_col_name}' created by combining '{col1}' and '{col2}'")

            # Update main lookup combo in step 2
            self.update_main_lookup_columns()

        except Exception as e:
            messagebox.showerror("Error", f"Error combining columns: {str(e)}")

    # Step 2 Methods
    def load_po_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PO File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.po_file_path = file_path
            self.po_file_label.config(text=Path(file_path).name, foreground="black")

            if file_path.endswith(('.xlsx', '.xls')):
                try:
                    xl_file = pd.ExcelFile(file_path)
                    sheets = xl_file.sheet_names
                    self.po_sheet_combo['values'] = sheets
                    self.po_sheet_combo.set(sheets[0])
                    self.po_sheet_combo['state'] = 'readonly'
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")
            else:
                self.load_po_sheet()

    def load_po_sheet(self):
        if not self.po_file_path:
            messagebox.showwarning("Warning", "Please select PO file first")
            return

        try:
            if self.po_file_path.endswith('.csv'):
                self.po_df = pd.read_csv(self.po_file_path)
            else:
                sheet_name = self.po_sheet_combo.get()
                self.po_df = pd.read_excel(self.po_file_path, sheet_name=sheet_name)

            columns = list(self.po_df.columns)
            self.po_lookup_combo['values'] = columns
            self.po_lookup_combo['state'] = 'readonly'

            # Create checkboxes for PO columns
            for widget in self.po_columns_frame.winfo_children():
                widget.destroy()
            self.po_column_vars.clear()

            for col in columns:
                var = tk.BooleanVar(value=False)
                self.po_column_vars[col] = var
                cb = ttk.Checkbutton(self.po_columns_frame, text=col, variable=var)
                cb.pack(anchor='w')

            messagebox.showinfo("Success", f"PO file loaded: {len(self.po_df)} rows")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading PO file: {str(e)}")

    def load_pr_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PR File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.pr_file_path = file_path
            self.pr_file_label.config(text=Path(file_path).name, foreground="black")

            if file_path.endswith(('.xlsx', '.xls')):
                try:
                    xl_file = pd.ExcelFile(file_path)
                    sheets = xl_file.sheet_names
                    self.pr_sheet_combo['values'] = sheets
                    self.pr_sheet_combo.set(sheets[0])
                    self.pr_sheet_combo['state'] = 'readonly'
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")
            else:
                self.load_pr_sheet()

    def load_pr_sheet(self):
        if not self.pr_file_path:
            messagebox.showwarning("Warning", "Please select PR file first")
            return

        try:
            if self.pr_file_path.endswith('.csv'):
                self.pr_df = pd.read_csv(self.pr_file_path)
            else:
                sheet_name = self.pr_sheet_combo.get()
                self.pr_df = pd.read_excel(self.pr_file_path, sheet_name=sheet_name)

            columns = list(self.pr_df.columns)
            self.pr_lookup_combo['values'] = columns
            self.pr_lookup_combo['state'] = 'readonly'

            # Create checkboxes for PR columns
            for widget in self.pr_columns_frame.winfo_children():
                widget.destroy()
            self.pr_column_vars.clear()

            for col in columns:
                var = tk.BooleanVar(value=False)
                self.pr_column_vars[col] = var
                cb = ttk.Checkbutton(self.pr_columns_frame, text=col, variable=var)
                cb.pack(anchor='w')

            messagebox.showinfo("Success", f"PR file loaded: {len(self.pr_df)} rows")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading PR file: {str(e)}")

    def update_main_lookup_columns(self):
        if self.main_df is not None:
            columns = list(self.main_df.columns)
            self.main_lookup_combo['values'] = columns
            self.main_lookup_combo['state'] = 'readonly'

    def perform_vlookup(self):
        if self.main_df is None:
            messagebox.showwarning("Warning", "Please load main file first")
            return

        if self.po_df is None or self.pr_df is None:
            messagebox.showwarning("Warning", "Please load both PO and PR files")
            return

        main_lookup_col = self.main_lookup_combo.get()
        po_lookup_col = self.po_lookup_combo.get()
        pr_lookup_col = self.pr_lookup_combo.get()

        if not main_lookup_col or not po_lookup_col or not pr_lookup_col:
            messagebox.showwarning("Warning", "Please select all lookup columns")
            return

        # Get selected columns
        po_cols = [col for col, var in self.po_column_vars.items() if var.get()]
        pr_cols = [col for col, var in self.pr_column_vars.items() if var.get()]

        if not po_cols and not pr_cols:
            messagebox.showwarning("Warning", "Please select at least one column from PO or PR")
            return

        try:
            # Prepare PO dataframe
            po_merge_cols = [po_lookup_col] + po_cols
            po_subset = self.po_df[po_merge_cols].copy()
            po_subset.columns = [po_lookup_col] + [f'PO {col}' for col in po_cols]

            # Prepare PR dataframe
            pr_merge_cols = [pr_lookup_col] + pr_cols
            pr_subset = self.pr_df[pr_merge_cols].copy()
            pr_subset.columns = [pr_lookup_col] + [f'PR {col}' for col in pr_cols]

            # Merge with main dataframe
            self.main_df = self.main_df.merge(po_subset,
                                             left_on=main_lookup_col,
                                             right_on=po_lookup_col,
                                             how='left')

            self.main_df = self.main_df.merge(pr_subset,
                                             left_on=main_lookup_col,
                                             right_on=pr_lookup_col,
                                             how='left')

            # Drop duplicate lookup columns if they were created
            if po_lookup_col != main_lookup_col and po_lookup_col in self.main_df.columns:
                self.main_df.drop(columns=[po_lookup_col], inplace=True)
            if pr_lookup_col != main_lookup_col and pr_lookup_col in self.main_df.columns:
                self.main_df.drop(columns=[pr_lookup_col], inplace=True)

            # Replace NaN, blanks, and #N/A with 0 for all PO and PR columns
            po_pr_columns = [f'PO {col}' for col in po_cols] + [f'PR {col}' for col in pr_cols]
            for col in po_pr_columns:
                if col in self.main_df.columns:
                    self.main_df[col] = self.main_df[col].fillna(0)
                    self.main_df[col] = self.main_df[col].replace('', 0)
                    self.main_df[col] = self.main_df[col].replace('#N/A', 0)

            self.step2_status.config(text="VLOOKUP completed successfully!")
            messagebox.showinfo("Success", f"VLOOKUP completed!\nAdded {len(po_cols)} PO columns and {len(pr_cols)} PR columns")

        except Exception as e:
            messagebox.showerror("Error", f"Error performing VLOOKUP: {str(e)}")

    # Step 3 Methods
    def add_month_combination(self):
        if self.main_df is None:
            messagebox.showwarning("Warning", "Please complete previous steps first")
            return

        row = len(self.month_combinations)

        frame = ttk.Frame(self.month_frame, relief='raised', borderwidth=1)
        frame.grid(row=row, column=0, sticky='ew', padx=5, pady=5)
        self.month_frame.columnconfigure(0, weight=1)

        # Get columns that start with 'PO ' or 'PR '
        po_pr_cols = [col for col in self.main_df.columns if col.startswith('PO ') or col.startswith('PR ')]

        ttk.Label(frame, text=f"Combination {row + 1}:").grid(row=0, column=0, sticky='w', padx=5, pady=5)

        ttk.Label(frame, text="PO Column:").grid(row=1, column=0, sticky='w', padx=5)
        po_combo = ttk.Combobox(frame, values=po_pr_cols, width=20)
        po_combo.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(frame, text="PR Column:").grid(row=2, column=0, sticky='w', padx=5)
        pr_combo = ttk.Combobox(frame, values=po_pr_cols, width=20)
        pr_combo.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(frame, text="Combined Name:").grid(row=3, column=0, sticky='w', padx=5)
        name_entry = ttk.Entry(frame, width=20)
        name_entry.grid(row=3, column=1, padx=5, pady=2)

        def remove_combination():
            frame.destroy()
            self.month_combinations.remove(combo_data)

        ttk.Button(frame, text="Remove", command=remove_combination).grid(row=4, column=0, columnspan=2, pady=5)

        combo_data = {
            'frame': frame,
            'po_combo': po_combo,
            'pr_combo': pr_combo,
            'name_entry': name_entry
        }
        self.month_combinations.append(combo_data)

    def create_combined_columns(self):
        if self.main_df is None:
            messagebox.showwarning("Warning", "Please complete previous steps first")
            return

        if not self.month_combinations:
            messagebox.showwarning("Warning", "Please add at least one month combination")
            return

        try:
            added_cols = []
            for combo in self.month_combinations:
                po_col = combo['po_combo'].get()
                pr_col = combo['pr_combo'].get()
                combined_name = combo['name_entry'].get()

                if not combined_name:
                    continue

                # If both columns selected, add them
                if po_col and pr_col and po_col in self.main_df.columns and pr_col in self.main_df.columns:
                    # Convert to numeric, replacing any non-numeric values with 0
                    po_values = pd.to_numeric(self.main_df[po_col], errors='coerce').fillna(0)
                    pr_values = pd.to_numeric(self.main_df[pr_col], errors='coerce').fillna(0)
                    self.main_df[combined_name] = po_values + pr_values
                    added_cols.append(combined_name)
                # If only one column, rename it
                elif po_col and po_col in self.main_df.columns:
                    self.main_df[combined_name] = pd.to_numeric(self.main_df[po_col], errors='coerce').fillna(0)
                    added_cols.append(combined_name)
                elif pr_col and pr_col in self.main_df.columns:
                    self.main_df[combined_name] = pd.to_numeric(self.main_df[pr_col], errors='coerce').fillna(0)
                    added_cols.append(combined_name)

            if added_cols:
                self.step3_status.config(text=f"Created {len(added_cols)} combined columns")
                messagebox.showinfo("Success", f"Created {len(added_cols)} combined columns:\n" + "\n".join(added_cols))
                self.update_preview()
            else:
                messagebox.showwarning("Warning", "No valid column combinations found")

        except Exception as e:
            messagebox.showerror("Error", f"Error creating combined columns: {str(e)}")

    # Step 4 Methods
    def update_preview(self):
        if self.main_df is None:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "No data available. Please complete previous steps.")
            return

        try:
            # Show first 10 rows
            preview = self.main_df.head(10).to_string()
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Preview (first 10 rows of {len(self.main_df)} total):\n\n{preview}\n\n")
            self.preview_text.insert(tk.END, f"Total columns: {len(self.main_df.columns)}\n")
            self.preview_text.insert(tk.END, f"Column names: {', '.join(self.main_df.columns)}")
        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Error generating preview: {str(e)}")

    def export_file(self):
        if self.main_df is None:
            messagebox.showwarning("Warning", "No data to export. Please complete previous steps.")
            return

        file_format = self.export_format.get()

        if file_format == "xlsx":
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Excel File"
            )
            if file_path:
                try:
                    self.main_df.to_excel(file_path, index=False)
                    self.step4_status.config(text=f"File exported successfully to {Path(file_path).name}")
                    messagebox.showinfo("Success", f"File exported successfully!\n{file_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"Error exporting file: {str(e)}")
        else:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Save CSV File"
            )
            if file_path:
                try:
                    self.main_df.to_csv(file_path, index=False)
                    self.step4_status.config(text=f"File exported successfully to {Path(file_path).name}")
                    messagebox.showinfo("Success", f"File exported successfully!\n{file_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"Error exporting file: {str(e)}")


# ==================== TOOL 3: SHORTAGE ALLOCATION ====================
class ShortageAllocationGUI:
    def __init__(self, parent):
        self.root = parent

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.item_code_col = tk.StringVar()
        self.output_path = tk.StringVar()
        self.df = None

        # --- Scrollable frame ---
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # --- GUI Widgets ---
        ttk.Label(self.scrollable_frame, text="Step 1: Select Excel File").pack(pady=5)
        file_frame = ttk.Frame(self.scrollable_frame)
        file_frame.pack(pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)

        ttk.Label(self.scrollable_frame, text="Step 2: Select Sheet").pack(pady=5)
        self.sheet_combo = ttk.Combobox(self.scrollable_frame, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.pack()

        ttk.Button(self.scrollable_frame, text="Load Columns", command=self.load_columns).pack(pady=10)

        ttk.Label(self.scrollable_frame, text="Step 3: Select Item Code Column").pack(pady=5)
        self.item_code_combo = ttk.Combobox(self.scrollable_frame, textvariable=self.item_code_col, state="readonly")
        self.item_code_combo.pack(pady=5)

        # --- PO Columns Frame ---
        self.po_frame = ttk.LabelFrame(self.scrollable_frame, text="PO Columns (Drag to Right to Select Priority)")
        self.po_frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        self.create_dual_listbox(self.po_frame, 'po')

        # --- Shortage Columns Frame ---
        self.short_frame = ttk.LabelFrame(self.scrollable_frame, text="Shortage Columns (Drag to Right to Select Priority)")
        self.short_frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        self.create_dual_listbox(self.short_frame, 'short')

        ttk.Button(self.scrollable_frame, text="Run Allocation", command=self.run_allocation).pack(pady=20)

    # --- Dual listbox creation ---
    def create_dual_listbox(self, parent, name):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)

        left = tk.Listbox(frame, selectmode=tk.SINGLE)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        right = tk.Listbox(frame)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Save references
        setattr(self, f'{name}_available', left)
        setattr(self, f'{name}_priority', right)

        # Bind double-click to move to priority
        left.bind('<Double-1>', lambda e, n=name: self.move_to_priority(n))

    # --- Move column to priority list ---
    def move_to_priority(self, name):
        available = getattr(self, f'{name}_available')
        priority = getattr(self, f'{name}_priority')
        idx = available.curselection()
        if not idx:
            return
        val = available.get(idx)
        if val not in priority.get(0, tk.END):
            priority.insert(tk.END, val)

    # --- File selection ---
    def browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.file_path.set(file)
            try:
                xls = pd.ExcelFile(file)
                self.sheet_combo['values'] = xls.sheet_names
                if xls.sheet_names:
                    self.sheet_combo.current(0)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read sheets: {e}")

    # --- Load columns ---
    def load_columns(self):
        try:
            self.df = pd.read_excel(self.file_path.get(), sheet_name=self.sheet_name.get())
            self.df.columns = self.df.columns.str.strip()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            return

        # Populate Item Code combobox
        self.item_code_combo['values'] = list(self.df.columns)
        if self.df.columns.any():
            self.item_code_combo.current(0)

        # Populate available columns for PO and Shortage
        for lst_name in ['po', 'short']:
            available = getattr(self, f'{lst_name}_available')
            available.delete(0, tk.END)
            priority = getattr(self, f'{lst_name}_priority')
            priority.delete(0, tk.END)
            for col in self.df.columns:
                available.insert(tk.END, col)

    # --- Run allocation ---
    def run_allocation(self):
        self.po_columns = list(self.po_priority.get(0, tk.END))
        self.shortage_columns = list(self.short_priority.get(0, tk.END))
        item_col = self.item_code_col.get()

        if not self.po_columns or not self.shortage_columns or not item_col:
            messagebox.showerror("Error", "Please select Item Code, at least one PO and one Shortage column.")
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx *.xls")])
        if not output_file:
            return
        self.output_path.set(output_file)

        df = self.df.copy()

        # === Allocation Logic (UNCHANGED) ===
        allocation_columns = [col + " ALLOCATED" for col in self.shortage_columns]
        for col in allocation_columns:
            df[col] = 0

        for po_col in self.po_columns:
            df[f'REMAINING {po_col}'] = df[po_col]

        for short_col, alloc_col in zip(self.shortage_columns, allocation_columns):
            for comp in df[item_col].unique():  # dynamic grouping
                comp_rows = df[df[item_col] == comp]
                total_short = comp_rows[short_col].max()
                if pd.isna(total_short) or total_short <= 0:
                    continue
                remaining = total_short
                for po_col in self.po_columns:
                    comp_sorted = comp_rows.sort_values(by=f'REMAINING {po_col}', ascending=False)
                    for idx in comp_sorted.index:
                        if remaining <= 0:
                            break
                        available = df.at[idx, f'REMAINING {po_col}']
                        if available > 0:
                            alloc_qty = min(available, remaining)
                            df.at[idx, alloc_col] += alloc_qty
                            df.at[idx, f'REMAINING {po_col}'] -= alloc_qty
                            remaining -= alloc_qty

        df['Total PO'] = df[self.po_columns].sum(axis=1)
        df['Total Shortage Requirement'] = df[self.shortage_columns].sum(axis=1)
        df['Total Allocated'] = df[allocation_columns].sum(axis=1)

        drop_cols = [col for col in df.columns if col.startswith("REMAINING ")] + ['Excess PO', 'Short PO']
        df_final = df.drop(columns=drop_cols, errors='ignore')

        df_final.to_excel(self.output_path.get(), sheet_name="Allocation", index=False)

        messagebox.showinfo("Success", f"Allocation complete.\nFile saved to: {self.output_path.get()}")

        print("\nOverall Summary Statistics (console only):")
        print(f"Total PO Available: {df['Total PO'].sum():,.0f}")
        print(f"Total Shortage Requirements: {df['Total Shortage Requirement'].sum():,.0f}")
        print(f"Total Allocated: {df['Total Allocated'].sum():,.0f}")


# ==================== MAIN FUNCTION ====================
def main():
    root = tk.Tk()
    app = MasterExcelToolkit(root)
    root.mainloop()


if __name__ == "__main__":
    main()

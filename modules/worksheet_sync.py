import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from typing import Dict, List
from templates.create_templates import create_worksheet_sync_template

class WorksheetSyncModule:
    def __init__(self, parent):
        self.parent = parent
        self.parent.configure(bg='#1a1a1a')

        self.workbook_path = None
        self.worksheets = []
        self.sync_data = {}

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#1a1a1a')
        style.configure('TLabel', background='#1a1a1a', foreground='#ffffff', font=('Segoe UI', 11))
        style.configure('TButton', font=('Segoe UI', 11, 'bold'), foreground='white', background='#008080', relief='flat')
        style.map('TButton', background=[('active', '#006666')])
        style.configure('Header.TLabel', font=('Segoe UI', 24, 'bold'), foreground='#ffffff')
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#b0b0b0')
        style.configure('Card.TFrame', background='#2d2d2d', relief='flat', borderwidth=1, bordercolor='#444444')
        style.configure("Treeview", background="#2d2d2d", foreground="white", fieldbackground="#2d2d2d", borderwidth=0, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", background="#008080", foreground="white", font=('Segoe UI', 11, 'bold'), relief='flat')
        style.map("Treeview.Heading", background=[('active', '#006666')])
        style.configure('TSpinbox', arrowsize=20)

    def create_widgets(self):
        main_frame = ttk.Frame(self.parent, padding=20)
        main_frame.pack(fill='both', expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        ttk.Label(header_frame, text="Worksheet Sync", style='Header.TLabel').pack(side='left')

        # --- Step 1: Workbook Selection ---
        selection_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        selection_card.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        selection_card.grid_columnconfigure(1, weight=1)
        ttk.Label(selection_card, text="Step 1: Select Workbook", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        self.file_path_var = tk.StringVar(value="No file selected")
        ttk.Button(selection_card, text="Browse Excel File", command=self.browse_file, width=25).grid(row=1, column=0, sticky="w", padx=(0, 10))
        ttk.Label(selection_card, textvariable=self.file_path_var, style='Status.TLabel').grid(row=1, column=1, sticky="ew")
        ttk.Button(selection_card, text="Generate Sample Workbook", command=self.generate_template, width=25).grid(row=2, column=0, sticky="w", pady=(10, 0))

        # --- Step 2: Configuration ---
        config_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        config_card.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        config_card.grid_columnconfigure(0, weight=1)
        config_card.grid_rowconfigure(3, weight=1) # Allow ws_frame row to expand
        config_card.grid_rowconfigure(3, weight=1) # Make the ws_frame row expandable
        ttk.Label(config_card, text="Step 2: Configure Sync", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        
        header_frame = ttk.Frame(config_card)
        header_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header_frame, text="Header Row:").pack(side='left', padx=(0, 10))
        self.header_var = tk.IntVar(value=1)
        ttk.Spinbox(header_frame, from_=1, to=100, textvariable=self.header_var, width=5).pack(side='left')

        ttk.Label(config_card, text="Select worksheets to sync:", font=('Segoe UI', 10, 'italic')).grid(row=2, column=0, columnspan=2, sticky='w', pady=(5,5))

        ws_frame = ttk.Frame(config_card)
        ws_frame.grid(row=3, column=0, sticky="nsew")
        ws_frame.grid_columnconfigure(0, weight=1)
        ws_frame.grid_rowconfigure(0, weight=1)
        self.ws_listbox = tk.Listbox(ws_frame, selectmode='multiple', background='#1e1e1e', foreground='white', borderwidth=0, highlightthickness=0, font=('Segoe UI', 10), height=5)
        self.ws_listbox.grid(row=0, column=0, sticky='nsew')
        ws_scrollbar = ttk.Scrollbar(ws_frame, orient='vertical', command=self.ws_listbox.yview)
        self.ws_listbox.configure(yscrollcommand=ws_scrollbar.set)
        ws_scrollbar.grid(row=0, column=1, sticky='ns')

        # --- Step 3: Preview, Sync & Export ---
        preview_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        preview_card.grid(row=3, column=0, sticky="nsew", pady=(0, 20))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_card.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        ttk.Label(preview_card, text="Step 3: Preview and Sync", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 15))

        tree_frame = ttk.Frame(preview_card)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.preview_tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        action_frame = ttk.Frame(preview_card)
        action_frame.grid(row=2, column=0, sticky="ew", pady=(15, 0))
        self.sync_btn = ttk.Button(action_frame, text="Sync Worksheets", command=self.sync_worksheets, state='disabled')
        self.sync_btn.pack(side='right', padx=(0, 10))
        self.export_btn = ttk.Button(action_frame, text="Export Synced Data", command=self.export_data, state='disabled')
        self.export_btn.pack(side='right')

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main_frame, textvariable=self.status_var, style='Status.TLabel').grid(row=4, column=0, sticky="ew", pady=(10, 0))

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel Workbook", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")] )
        if file_path:
            self.workbook_path = file_path
            self.file_path_var.set(os.path.basename(file_path))
            self.load_workbook()

    def load_workbook(self):
        try:
            self.status_var.set(f"Loading {os.path.basename(self.workbook_path)}...")
            self.worksheets = pd.ExcelFile(self.workbook_path).sheet_names
            self.ws_listbox.delete(0, 'end')
            for ws in self.worksheets:
                self.ws_listbox.insert('end', ws)
            self.sync_btn.config(state='normal')
            self.status_var.set("Workbook loaded. Select worksheets to sync.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workbook: {str(e)}")
            self.status_var.set("Error loading workbook.")

    def sync_worksheets(self):
        selected_indices = self.ws_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one worksheet to sync.")
            return

        selected_worksheets = [self.ws_listbox.get(i) for i in selected_indices]
        try:
            self.status_var.set("Syncing worksheets...")
            self.sync_data = None
            dataframes = []
            header_row_index = self.header_var.get() - 1

            for ws_name in selected_worksheets:
                df = pd.read_excel(self.workbook_path, sheet_name=ws_name, header=header_row_index)
                df['Source Worksheet'] = ws_name
                dataframes.append(df)

            if dataframes:
                merged_df = pd.concat(dataframes, ignore_index=True)
                # Reorder columns to have 'Source Worksheet' first
                cols = ['Source Worksheet'] + [col for col in merged_df.columns if col != 'Source Worksheet']
                self.sync_data = merged_df[cols]

            if self.sync_data is not None and not self.sync_data.empty:
                self.update_synced_preview()
                self.status_var.set("Worksheets synced successfully.")
                self.export_btn.config(state='normal')
            else:
                self.status_var.set("No data found in selected worksheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to sync worksheets: {str(e)}")
            self.status_var.set("Error during worksheet sync.")

    def update_synced_preview(self):
        if self.sync_data is None:
            return

        self.preview_tree.delete(*self.preview_tree.get_children())
        
        headers = self.sync_data.columns.tolist()
        self.preview_tree['columns'] = headers
        for col in headers:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=120, minwidth=60)

        for index, row in self.sync_data.iterrows():
            self.preview_tree.insert('', 'end', values=row.tolist())

    def generate_template(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save Sample Workbook")
        if not file_path:
            return
        try:
            sheets = create_worksheet_sync_template()
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Sample workbook generated successfully at {file_path}")
            self.workbook_path = file_path
            self.file_path_var.set(os.path.basename(file_path))
            self.load_workbook()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate template: {str(e)}")

    def export_data(self):
        if self.sync_data is None:
            messagebox.showwarning("Warning", "No synced data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save Synced Data")
        if not file_path:
            return
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                self.sync_data.to_excel(writer, sheet_name='Synced_Data', index=False)
            messagebox.showinfo("Success", f"Data exported successfully to {file_path}")
            self.status_var.set("Data exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
            self.status_var.set("Error exporting data.")
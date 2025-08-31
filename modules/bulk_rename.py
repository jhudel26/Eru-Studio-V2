import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import glob
import pandas as pd
import os
import shutil
from pathlib import Path
from typing import List, Dict

# Import the template creation function

class BulkRenameModule:
    def __init__(self, parent):
        self.parent = parent
        self.parent.configure(bg='#1a1a1a')  # Set background color for the module window

        self.template_path = None
        self.template_data = None
        self.source_folder = None
        self.rename_mapping = {}
        self.preview_data = []
        self.search_var = tk.StringVar()

        self.setup_styles()
        self.create_widgets()
        self.search_var.trace_add("write", self.search_files)

    def setup_styles(self):
        """Configure custom styles for the module"""
        style = ttk.Style()
        style.theme_use('clam')

        # General widget styles
        style.configure('TFrame', background='#1a1a1a')
        style.configure('TLabel', background='#1a1a1a', foreground='#ffffff', font=('Segoe UI', 11))
        style.configure('TButton', font=('Segoe UI', 11, 'bold'), foreground='white', background='#008080', relief='flat')
        style.map('TButton', background=[('active', '#006666')])
        style.configure('Header.TLabel', font=('Segoe UI', 24, 'bold'), foreground='#ffffff')
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#b0b0b0')
        
        # Combobox style
        style.configure('TCombobox', fieldbackground='#2d2d2d', background='#2d2d2d', foreground='#ffffff', 
                        arrowcolor='#ffffff', bordercolor='#4A90E2', lightcolor='#2d2d2d', darkcolor='#2d2d2d',
                        selectbackground='#357ABD', selectforeground='white')
        
        # Treeview style
        style.configure("Treeview",
                        background="#2d2d2d",
                        foreground="white",
                        fieldbackground="#2d2d2d",
                        borderwidth=0,
                        font=('Segoe UI', 10))
        style.configure("Treeview.Heading",
                        background="#008080",
                        foreground="white",
                        font=('Segoe UI', 11, 'bold'),
                        relief='flat')
        style.map("Treeview.Heading", background=[('active', '#006666')])
        
        # Custom frame style
        style.configure('Card.TFrame', background='#2d2d2d', relief='flat', borderwidth=1, bordercolor='#444444')

    def create_widgets(self):
        """Create the modern module interface"""
        main_frame = ttk.Frame(self.parent, padding=20)
        main_frame.pack(fill='both', expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        ttk.Label(header_frame, text="Bulk Rename", style='Header.TLabel').pack(side='left')

        # --- Step 1: Inputs ---
        inputs_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        inputs_card.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        inputs_card.grid_columnconfigure(1, weight=1)
        ttk.Label(inputs_card, text="Step 1: Select Inputs", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))

        # Template selection
        self.template_path_var = tk.StringVar(value="No template selected")
        ttk.Button(inputs_card, text="Browse Excel Template", command=self.browse_template, width=25).grid(row=1, column=0, sticky="w", padx=(0, 10))
        ttk.Label(inputs_card, textvariable=self.template_path_var, style='Status.TLabel').grid(row=1, column=1, columnspan=2, sticky="ew")
        ttk.Button(inputs_card, text="Generate Template", command=self.generate_template, width=25).grid(row=2, column=0, sticky="w", pady=(10, 0))

        # Source folder selection
        self.source_path_var = tk.StringVar(value="No source folder selected")
        ttk.Button(inputs_card, text="Browse Source Folder", command=self.browse_source_folder, width=25).grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(inputs_card, textvariable=self.source_path_var, style='Status.TLabel').grid(row=3, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        # --- Step 2: Configuration ---
        config_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        config_card.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        config_card.grid_columnconfigure(1, weight=1)
        ttk.Label(config_card, text="Step 2: Configure Columns", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 15))

        ttk.Label(config_card, text="Current Filename Column:").grid(row=1, column=0, sticky="w", padx=(0, 5))
        self.current_col_var = tk.StringVar()
        self.current_col_combo = ttk.Combobox(config_card, textvariable=self.current_col_var, state='readonly', width=25)
        self.current_col_combo.grid(row=1, column=1, sticky="w", padx=(0, 20))

        ttk.Label(config_card, text="New Filename Column:").grid(row=1, column=2, sticky="w", padx=(0, 5))
        self.new_col_var = tk.StringVar()
        self.new_col_combo = ttk.Combobox(config_card, textvariable=self.new_col_var, state='readonly', width=25)
        self.new_col_combo.grid(row=1, column=3, sticky="w")

        # --- Step 3: Preview & Execute ---
        preview_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        preview_card.grid(row=3, column=0, sticky="nsew", pady=(0, 20))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_card.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        ttk.Label(preview_card, text="Step 3: Preview and Execute", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

        # Search Bar
        search_frame = ttk.Frame(preview_card)
        search_frame.grid(row=1, column=0, sticky="ew", pady=(5, 10))
        search_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, sticky='w', padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky='ew')

        # Treeview
        tree_frame = ttk.Frame(preview_card)
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ('current_name', 'new_name', 'status')
        self.preview_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        self.preview_tree.heading('current_name', text='Current Filename')
        self.preview_tree.heading('new_name', text='New Filename')
        self.preview_tree.heading('status', text='Status')
        self.preview_tree.column('current_name', width=250)
        self.preview_tree.column('new_name', width=250)
        self.preview_tree.column('status', width=150)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky='ns')

        self.preview_tree.tag_configure('ready', foreground='#73d16e')
        self.preview_tree.tag_configure('skip', foreground='#f0e68c')
        self.preview_tree.tag_configure('error', foreground='#ff6b6b')

        # Action buttons
        action_frame = ttk.Frame(preview_card)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        action_frame.grid_columnconfigure(1, weight=1)

        self.create_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(action_frame, text="Create backup before renaming", variable=self.create_backup).grid(row=0, column=0, sticky='w')

        button_group = ttk.Frame(action_frame)
        button_group.grid(row=0, column=2, sticky='e')

        self.preview_btn = ttk.Button(button_group, text="Generate Preview", command=self.generate_preview, width=20)
        self.preview_btn.pack(side='right', padx=(0, 10))

        self.rename_btn = ttk.Button(button_group, text="Rename Files", command=self.execute_rename, state='disabled', width=20)
        self.rename_btn.pack(side='right', padx=(0, 10))

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, style='Status.TLabel')
        status_label.grid(row=4, column=0, sticky="ew", pady=(10, 0))

    def generate_template(self):
        """Generate an Excel template pre-filled with filenames from the source directory."""
        if not self.source_folder:
            messagebox.showwarning("Warning", "Please select a source folder first.")
            return

        try:
            all_paths = glob.glob(os.path.join(self.source_folder, '*'))
            items = [os.path.basename(path) for path in all_paths if not os.path.basename(path).startswith('backup_')]
            if not items:
                messagebox.showinfo("Info", "No items found in the source directory.")
                return

            df = pd.DataFrame({'Current Name': items, 'New Name': ''})

            save_path = filedialog.asksaveasfilename(
                title="Save Template",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile="bulk_rename_template.xlsx"
            )

            if not save_path:
                self.status_var.set("Template generation cancelled.")
                return

            df.to_excel(save_path, index=False, sheet_name='Rename_Mapping')
            messagebox.showinfo("Success", f"Template with {len(items)} item(s) saved to:\n{save_path}")
            self.status_var.set("Template generated successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate template: {e}")
            self.status_var.set("Error generating template.")

    def browse_template(self):
        """Browse and select Excel template file"""
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.template_path = path
            self.template_path_var.set(os.path.basename(path))
            self.load_template()
            self.status_var.set("Template loaded. Select source folder.")

    def browse_source_folder(self):
        """Browse and select source folder"""
        path = filedialog.askdirectory()
        if path:
            self.source_folder = path
            self.source_path_var.set(path)
            if self.template_data is not None:
                self.preview_btn.config(state='normal')
            self.status_var.set("Source folder selected. Ready to generate preview.")

    def load_template(self):
        """Load the Excel template and populate column dropdowns"""
        if not self.template_path:
            return
        try:
            self.template_data = pd.read_excel(self.template_path)
            columns = self.template_data.columns.tolist()
            self.current_col_combo['values'] = columns
            self.new_col_combo['values'] = columns
            
            # Auto-select columns if they match expected names
            if 'Current Name' in columns:
                self.current_col_var.set('Current Name')
            if 'New Name' in columns:
                self.new_col_var.set('New Name')
                
            if self.source_folder:
                self.preview_btn.config(state='normal')
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load template: {str(e)}")
            self.status_var.set("Error loading template.")

    def generate_preview(self):
        """Generate preview of rename operations"""
        if not all([self.source_folder, self.template_data is not None, self.current_col_var.get(), self.new_col_var.get()]):
            messagebox.showwarning("Warning", "Please select template, source folder, and columns.")
            return

        try:
            self.status_var.set("Generating preview...")
            self.preview_tree.delete(*self.preview_tree.get_children())
            self.preview_data = []
            self.rename_btn.config(state='disabled')

            current_col = self.current_col_var.get()
            new_col = self.new_col_var.get()
            
            self.rename_mapping = dict(zip(self.template_data[current_col].astype(str), self.template_data[new_col].astype(str)))

            all_paths = glob.glob(os.path.join(self.source_folder, '*'))
            source_items = [os.path.basename(path) for path in all_paths if not os.path.basename(path).startswith('backup_')]

            for filename in source_items:
                if filename in self.rename_mapping:
                    new_name_from_template = self.rename_mapping[filename]

                    # Handle cases where 'New Name' in Excel is blank or just whitespace
                    if pd.isna(new_name_from_template) or str(new_name_from_template).strip() in ['', 'nan']:
                        self.preview_tree.insert('', 'end', values=(filename, "No change", "ℹ️ New name is empty"), tags=('skip',))
                        continue

                    new_name_str = str(new_name_from_template).strip()
                    
                    # If the new name from the template doesn't have an extension, reuse the old one
                    _ , old_ext = os.path.splitext(filename)
                    _ , new_ext = os.path.splitext(new_name_str)

                    if not new_ext and old_ext:
                        final_new_name = f"{new_name_str}{old_ext}"
                    else:
                        final_new_name = new_name_str

                    status = "✅ Ready"
                    self.preview_tree.insert('', 'end', values=(filename, final_new_name, status), tags=('ready',))
                    self.preview_data.append({'current': filename, 'new': final_new_name, 'status': status})
                else:
                    self.preview_tree.insert('', 'end', values=(filename, "No change", "ℹ️ Not in template"), tags=('skip',))
            
            if any(item['status'] == "✅ Ready" for item in self.preview_data):
                self.rename_btn.config(state='normal')
                
            self.status_var.set(f"Preview generated: {len(source_items)} items checked.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            self.status_var.set("Error generating preview.")

    def search_files(self, *args):
        query = self.search_var.get().lower()
        self.preview_tree.delete(*self.preview_tree.get_children())

        if not self.preview_data:
            return

        for item in self.preview_data:
            current_name = item.get('current', '')
            new_name = item.get('new', 'No change')
            status = item.get('status', '')
            
            if query in current_name.lower() or query in new_name.lower():
                tag = ''
                if status == "✅ Ready":
                    tag = 'ready'
                elif status == "ℹ️ Not in template":
                    tag = 'skip'
                elif "Error" in status or "Exists" in status:
                    tag = 'error'
                
                self.preview_tree.insert('', 'end', values=(current_name, new_name, status), tags=(tag,))

    def execute_rename(self):
        """Execute the rename operations"""
        ready_items = [item for item in self.preview_data if item['status'] == "✅ Ready"]
        if not ready_items:
            messagebox.showinfo("Info", "No files are ready for renaming.")
            return

        if not messagebox.askyesno("Confirm Rename", f"Are you sure you want to rename {len(ready_items)} files?"):
            return

        try:
            self.status_var.set("Renaming files...")
            renamed_count = 0
            errors = []

            if self.create_backup.get():
                backup_folder = os.path.join(self.source_folder, f"backup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}")
                os.makedirs(backup_folder, exist_ok=True)
                for item in ready_items:
                    source_path = os.path.join(self.source_folder, item['current'])
                    dest_path = os.path.join(backup_folder, item['current'])
                    if os.path.isdir(source_path):
                        shutil.copytree(source_path, dest_path)
                    else:
                        shutil.copy2(source_path, backup_folder)

            for item in ready_items:
                try:
                    current_path = os.path.join(self.source_folder, item['current'])
                    new_path = os.path.join(self.source_folder, item['new'])
                    
                    if os.path.exists(new_path):
                        raise FileExistsError(f"Target file '{item['new']}' already exists.")

                    os.rename(current_path, new_path)
                    renamed_count += 1
                    item['status'] = "✅ Renamed"
                except Exception as e:
                    errors.append(f"{item['current']} -> {item['new']}: {e}")
                    item['status'] = "❌ Error"
            
            self.generate_preview() # Refresh the preview
            
            if errors:
                error_details = "\n".join(errors)
                messagebox.showwarning("Rename Complete with Errors", f"Renamed {renamed_count} files.\n\nErrors occurred:\n{error_details}")
            else:
                messagebox.showinfo("Success", f"Successfully renamed {renamed_count} files!")

            self.status_var.set(f"Rename complete. {renamed_count} files processed.")
            self.rename_btn.config(state='disabled')

        except Exception as e:
            messagebox.showerror("Error", f"A critical error occurred during renaming: {e}")
            self.status_var.set("Error during rename execution.")
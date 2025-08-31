import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from typing import Dict

# Import the template creation function
from templates.create_templates import create_folder_creator_template

class FolderCreatorModule:
    def __init__(self, parent):
        self.parent = parent
        self.parent.configure(bg='#1a1a1a')

        self.template_path = None
        self.template_data = None
        self.output_folder = None
        self.folder_structure = []
        self.search_var = tk.StringVar()

        self.setup_styles()
        self.create_widgets()
        self.search_var.trace_add("write", self.search_folders)

    def setup_styles(self):
        """Configure custom styles for the module"""
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('TFrame', background='#1a1a1a')
        style.configure('TLabel', background='#1a1a1a', foreground='#ffffff', font=('Segoe UI', 11))
        style.configure('TButton', font=('Segoe UI', 11, 'bold'), foreground='white', background='#008080', relief='flat')
        style.map('TButton', background=[('active', '#006666')])
        style.configure('Header.TLabel', font=('Segoe UI', 24, 'bold'), foreground='#ffffff')
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#b0b0b0')
        
        style.configure('TCombobox', fieldbackground='#2d2d2d', background='#2d2d2d', foreground='#ffffff', 
                        arrowcolor='#ffffff', bordercolor='#4A90E2', lightcolor='#2d2d2d', darkcolor='#2d2d2d',
                        selectbackground='#357ABD', selectforeground='white')
        
        style.configure("Treeview", background="#2d2d2d", foreground="white", fieldbackground="#2d2d2d",
                        borderwidth=0, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", background="#008080", foreground="white", 
                        font=('Segoe UI', 11, 'bold'), relief='flat')
        style.map("Treeview.Heading", background=[('active', '#006666')])
        
        style.configure('Card.TFrame', background='#2d2d2d', relief='flat', borderwidth=1, bordercolor='#444444')

    def create_widgets(self):
        """Create the modern module interface"""
        main_frame = ttk.Frame(self.parent, padding=20)
        main_frame.pack(fill='both', expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        ttk.Label(header_frame, text="Folder Creator", style='Header.TLabel').pack(side='left')

        # --- Step 1: Template ---
        template_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        template_card.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        template_card.grid_columnconfigure(0, weight=1)
        ttk.Label(template_card, text="Step 1: Select or Generate Template", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        self.template_path_var = tk.StringVar(value="No template selected")
        ttk.Label(template_card, textvariable=self.template_path_var, style='Status.TLabel').grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))
        button_frame1 = ttk.Frame(template_card)
        button_frame1.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Button(button_frame1, text="Browse Excel Template", command=self.browse_template, width=25).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame1, text="Generate Template", command=self.generate_template, width=25).pack(side='left')

        # --- Step 2: Output Location ---
        source_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        source_card.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        source_card.grid_columnconfigure(0, weight=1)
        ttk.Label(source_card, text="Step 2: Select Output Location", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 15))
        self.output_path_var = tk.StringVar(value="No folder selected")
        ttk.Label(source_card, textvariable=self.output_path_var, style='Status.TLabel').grid(row=1, column=0, sticky="w", pady=(0, 10))
        ttk.Button(source_card, text="Browse Output Folder", command=self.browse_output_folder, width=25).grid(row=2, column=0, sticky="w")

        # --- Step 3: Configuration & Preview ---
        preview_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        preview_card.grid(row=3, column=0, sticky="nsew", pady=(0, 20))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_card.grid_rowconfigure(3, weight=1)
        ttk.Label(preview_card, text="Step 3: Configure and Preview", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

        # Configuration row
        col_frame = ttk.Frame(preview_card)
        col_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(col_frame, text="Folder Name Column:").pack(side='left', padx=(0, 10))
        self.folder_name_col_var = tk.StringVar()
        self.folder_name_combo = ttk.Combobox(col_frame, textvariable=self.folder_name_col_var, state='readonly', width=25)
        self.folder_name_combo.pack(side='left', padx=(0, 20))
        ttk.Label(col_frame, text="Parent Folder Column:").pack(side='left', padx=(0, 10))
        self.parent_folder_col_var = tk.StringVar()
        self.parent_folder_combo = ttk.Combobox(col_frame, textvariable=self.parent_folder_col_var, state='readonly', width=25)
        self.parent_folder_combo.pack(side='left')

        # Search row
        search_frame = ttk.Frame(preview_card)
        search_frame.grid(row=2, column=0, sticky='ew', pady=(5, 10))
        search_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="Search:", font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w', padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky='ew')

        # Preview tree in a frame to manage scrollbar
        tree_frame = ttk.Frame(preview_card)
        tree_frame.grid(row=3, column=0, sticky='nsew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        columns = ('name', 'parent', 'path', 'status')
        self.preview_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        self.preview_tree.heading('name', text='Folder Name')
        self.preview_tree.heading('parent', text='Parent Folder')
        self.preview_tree.heading('path', text='Relative Path')
        self.preview_tree.heading('status', text='Status')
        self.preview_tree.column('name', width=200)
        self.preview_tree.column('parent', width=150)
        self.preview_tree.column('path', width=300)
        self.preview_tree.column('status', width=120)

        # Action buttons
        action_frame = ttk.Frame(preview_card)
        action_frame.grid(row=4, column=0, sticky='ew', pady=(15, 0))
        action_frame.grid_columnconfigure(1, weight=1)

        self.create_readme = tk.BooleanVar(value=False)
        ttk.Checkbutton(action_frame, text="Create README.txt in each folder", variable=self.create_readme).grid(row=0, column=0, sticky='w')

        button_group = ttk.Frame(action_frame)
        button_group.grid(row=0, column=2, sticky='e')

        self.preview_btn = ttk.Button(button_group, text="Generate Preview", command=self.generate_preview, state='disabled', width=20)
        self.preview_btn.pack(side='right', padx=(0, 10))

        self.create_btn = ttk.Button(button_group, text="Create Folders", command=self.create_folders, state='disabled', width=20)
        self.create_btn.pack(side='right')

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main_frame, textvariable=self.status_var, style='Status.TLabel').grid(row=5, column=0, sticky="ew", pady=(10, 0))

    def generate_template(self):
        """Generate a sample Excel template for folder creation."""
        try:
            df = create_folder_creator_template()
            save_path = filedialog.asksaveasfilename(
                title="Save Sample Template",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile="folder_creator_template.xlsx"
            )
            if not save_path:
                self.status_var.set("Template generation cancelled.")
                return
            df.to_excel(save_path, index=False, sheet_name='Folder_Structure')
            messagebox.showinfo("Success", f"Sample template saved to:\n{save_path}")
            self.status_var.set("Sample template generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate template: {e}")
            self.status_var.set("Error generating template.")

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.template_path = path
            self.template_path_var.set(os.path.basename(path))
            self.load_template()
            self.status_var.set("Template loaded. Select output location.")

    def browse_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.output_path_var.set(path)
            if self.template_data is not None:
                self.preview_btn.config(state='normal')
            self.status_var.set("Output location selected. Ready to generate preview.")

    def load_template(self):
        if not self.template_path:
            return
        try:
            self.template_data = pd.read_excel(self.template_path, na_filter=False)
            columns = self.template_data.columns.tolist()
            self.folder_name_combo['values'] = columns
            self.parent_folder_combo['values'] = [''] + columns
            if 'Folder Name' in columns:
                self.folder_name_col_var.set('Folder Name')
            if 'Parent Folder' in columns:
                self.parent_folder_col_var.set('Parent Folder')
            if self.output_folder:
                self.preview_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load template: {str(e)}")
            self.status_var.set("Error loading template.")

    def generate_preview(self):
        if not all([self.output_folder, self.template_data is not None, self.folder_name_col_var.get()]):
            messagebox.showwarning("Warning", "Please select template, output folder, and folder name column.")
            return
        try:
            self.status_var.set("Generating preview...")
            self.preview_tree.delete(*self.preview_tree.get_children())
            self.folder_structure = []
            self.create_btn.config(state='disabled')

            folder_col = self.folder_name_col_var.get()
            parent_col = self.parent_folder_col_var.get()

            for _, row in self.template_data.iterrows():
                folder_name = str(row[folder_col]).strip()
                if not folder_name:
                    continue
                
                parent_name = str(row[parent_col]).strip() if parent_col and parent_col in row and pd.notna(row[parent_col]) and str(row[parent_col]).strip() else ""
                
                full_path = os.path.join(self.output_folder, parent_name, folder_name) if parent_name else os.path.join(self.output_folder, folder_name)
                
                status = "⚠️ Exists" if os.path.exists(full_path) else "✅ Ready"
                
                item = {'name': folder_name, 'parent': parent_name, 'full_path': full_path, 'status': status}
                self.folder_structure.append(item)
                self.preview_tree.insert('', 'end', values=(folder_name, parent_name if parent_name else "<ROOT>", os.path.relpath(full_path, self.output_folder), status))

            if any(item['status'] == "✅ Ready" for item in self.folder_structure):
                self.create_btn.config(state='normal')
            self.status_var.set(f"Preview generated: {len(self.folder_structure)} folders planned.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            self.status_var.set("Error generating preview.")

    def search_folders(self, *args):
        query = self.search_var.get().lower()
        self.preview_tree.delete(*self.preview_tree.get_children())

        if not self.folder_structure:
            return

        if not query:
            # If search is cleared, show all items from the original preview
            for item in self.folder_structure:
                self.preview_tree.insert('', 'end', values=(
                    item['name'], 
                    item['parent'] if item['parent'] else "<ROOT>", 
                    os.path.relpath(item['full_path'], self.output_folder),
                    item['status']
                ))
        else:
            # Filter and show only matching items
            for item in self.folder_structure:
                if query in item['name'].lower():
                    self.preview_tree.insert('', 'end', values=(
                        item['name'], 
                        item['parent'] if item['parent'] else "<ROOT>", 
                        os.path.relpath(item['full_path'], self.output_folder),
                        item['status']
                    ))

    def create_folders(self):
        ready_items = [item for item in self.folder_structure if item['status'] == "✅ Ready"]
        if not ready_items:
            messagebox.showinfo("Info", "No new folders to create.")
            return
        if not messagebox.askyesno("Confirm Creation", f"Are you sure you want to create {len(ready_items)} folders?"):
            return
        try:
            self.status_var.set("Creating folders...")
            created_count = 0
            errors = []

            for item in ready_items:
                try:
                    os.makedirs(item['full_path'], exist_ok=True)
                    if self.create_readme.get():
                        with open(os.path.join(item['full_path'], "README.txt"), 'w', encoding='utf-8') as f:
                            f.write(f"Folder created by EruStudio.\nName: {item['name']}\nParent: {item['parent'] or 'N/A'}")
                    created_count += 1
                    item['status'] = "✅ Created"
                except Exception as e:
                    errors.append(f"{item['name']}: {e}")
                    item['status'] = "❌ Error"
            
            self.generate_preview() # Refresh preview

            if errors:
                messagebox.showwarning("Creation Complete with Errors", f"Created {created_count} folders.\n\nErrors:\n" + "\n".join(errors))
            else:
                messagebox.showinfo("Success", f"Successfully created {created_count} folders!")
            
            self.status_var.set(f"Creation complete. {created_count} folders processed.")
            self.create_btn.config(state='disabled')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create folders: {str(e)}")
            self.status_var.set("Error creating folders.")
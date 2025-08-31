import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import zipfile
import threading
from typing import Dict
import pandas as pd

# Import the template creation function

class MultiZipModule:
    def __init__(self, parent):
        self.parent = parent
        self.parent.configure(bg='#1a1a1a')

        self.source_folder = None
        self.output_folder = None
        self.template_path = None
        self.zip_tasks = []
        self.search_var = tk.StringVar()
        self.template_path_var = tk.StringVar(value="No template selected")

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
        
        style.configure("Treeview", background="#2d2d2d", foreground="white", fieldbackground="#2d2d2d",
                        borderwidth=0, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", background="#008080", foreground="white", 
                        font=('Segoe UI', 11, 'bold'), relief='flat')
        style.map("Treeview.Heading", background=[('active', '#006666')])
        
        style.configure('Card.TFrame', background='#2d2d2d', relief='flat', borderwidth=1, bordercolor='#444444')
        style.configure('TCheckbutton', background='#2d2d2d', foreground='#ffffff', font=('Segoe UI', 10))
        style.map('TCheckbutton', background=[('active', '#2d2d2d')])

    def create_widgets(self):
        """Create the modern module interface"""
        main_frame = ttk.Frame(self.parent, padding=20)
        main_frame.pack(fill='both', expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 30))
        ttk.Label(header_frame, text="Multi-Zip", style='Header.TLabel').pack(side='left')
        ttk.Button(header_frame, text="Generate Template", command=self.generate_template, width=25).pack(side='right')

        # --- Step 1: Source & Output ---
        io_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        io_card.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        io_card.grid_columnconfigure(1, weight=1)
        ttk.Label(io_card, text="Step 1: Select Source and Output", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        
        ttk.Button(io_card, text="Browse Source Folder", command=self.browse_source_folder, width=25).grid(row=1, column=0, sticky="w", padx=(0, 10))
        self.source_path_var = tk.StringVar(value="No source folder selected")
        ttk.Label(io_card, textvariable=self.source_path_var, style='Status.TLabel').grid(row=1, column=1, sticky="ew")

        ttk.Button(io_card, text="Browse Output Folder", command=self.browse_output_folder, width=25).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(10,0))
        self.output_path_var = tk.StringVar(value="No output folder selected")
        ttk.Label(io_card, textvariable=self.output_path_var, style='Status.TLabel').grid(row=2, column=1, sticky="ew", pady=(10,0))

        ttk.Button(io_card, text="Browse Template", command=self.browse_template_file, width=25).grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(10,0))
        ttk.Label(io_card, textvariable=self.template_path_var, style='Status.TLabel').grid(row=3, column=1, sticky="ew", pady=(10,0))

        # --- Step 2: Configuration & Preview ---
        preview_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        preview_card.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_card.grid_rowconfigure(3, weight=1) # Adjust row for treeview
        main_frame.grid_rowconfigure(2, weight=1)
        ttk.Label(preview_card, text="Step 2: Configure and Preview", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        
        options_frame = ttk.Frame(preview_card)
        options_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self.zip_top_level_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Zip top-level folders only", variable=self.zip_top_level_only, command=self.scan_folders).pack(side='left', padx=(0, 20))
        self.include_root_dir = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Include root folder in archive", variable=self.include_root_dir).pack(side='left', padx=(0, 20))

        # --- Search Bar ---
        search_frame = ttk.Frame(preview_card)
        search_frame.grid(row=2, column=0, sticky="ew", pady=(10, 10))
        search_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="Search:", font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w', padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky='ew')

        # --- Treeview for Preview ---     
        tree_frame = ttk.Frame(preview_card)
        tree_frame.grid(row=3, column=0, columnspan=2, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        columns = ('name', 'path', 'size', 'status')
        self.preview_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        self.preview_tree.heading('name', text='Folder Name')
        self.preview_tree.heading('path', text='Relative Path')
        self.preview_tree.heading('size', text='Size')
        self.preview_tree.heading('status', text='Status')
        self.preview_tree.column('name', width=200)
        self.preview_tree.column('path', width=250)
        self.preview_tree.column('size', width=100, anchor='e')
        self.preview_tree.column('status', width=120)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky='ns')

        # --- Step 3: Execution ---
        action_card = ttk.Frame(main_frame, style='Card.TFrame', padding=20)
        action_card.grid(row=3, column=0, sticky="ew")
        action_card.grid_columnconfigure(1, weight=1)
        ttk.Label(action_card, text="Step 3: Execute Zipping", font=('Segoe UI', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        self.zip_btn = ttk.Button(action_card, text="Create Zip Archives", command=self.create_zip_files, state='disabled', width=25)
        self.zip_btn.grid(row=1, column=0, sticky="w")
        self.progress_bar = ttk.Progressbar(action_card, orient='horizontal', mode='determinate')
        self.progress_bar.grid(row=1, column=1, sticky="ew", padx=(20, 0))

        # Status bar
        self.status_var = tk.StringVar(value="Ready. Select source and output folders.")
        ttk.Label(main_frame, textvariable=self.status_var, style='Status.TLabel').grid(row=4, column=0, sticky="ew", pady=(10, 0))

    def generate_template(self):
        """Generate an Excel template pre-filled with folder names from the source directory."""
        if not self.source_folder:
            messagebox.showwarning("Warning", "Please select a source folder first.")
            return
        try:
            folder_names = [d for d in os.listdir(self.source_folder) if os.path.isdir(os.path.join(self.source_folder, d))]
            if not folder_names:
                messagebox.showinfo("Info", "No folders found in the source directory.")
                return

            df = pd.DataFrame({'Folder Name': folder_names})
            
            save_path = filedialog.asksaveasfilename(
                title="Save Template",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile="multi_zip_template.xlsx"
            )
            if not save_path:
                self.status_var.set("Template generation cancelled.")
                return
            
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Template with {len(folder_names)} folder(s) saved to:\n{save_path}")
            self.status_var.set("Template generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate template: {e}")
            self.status_var.set("Error generating template.")

    def browse_template_file(self):
        path = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if path:
            self.template_path = path
            self.template_path_var.set(os.path.basename(path))
            self.scan_folders()

    def browse_source_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.source_folder = path
            self.source_path_var.set(path)
            self.scan_folders()

    def browse_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.output_path_var.set(path)
            self.scan_folders()

    def scan_folders(self):
        if not all([self.source_folder, self.output_folder]):
            self.status_var.set("Please select both source and output folders first.")
            return

        self.status_var.set("Scanning folders...")
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.zip_tasks = []
        self.zip_btn.config(state='disabled')

        try:
            folders_to_process = []
            base_path = self.source_folder

            if self.template_path:
                df = pd.read_excel(self.template_path)
                if 'Folder Name' not in df.columns:
                    messagebox.showerror("Error", "Template must have a 'Folder Name' column.")
                    self.status_var.set("Error: Invalid template.")
                    return
                
                for folder_name in df['Folder Name'].dropna().astype(str):
                    full_path = os.path.join(self.source_folder, folder_name)
                    if os.path.isdir(full_path):
                        folders_to_process.append({'name': folder_name, 'path': full_path})
            else: # Fallback to original behavior
                if self.zip_top_level_only.get():
                    for d in os.listdir(self.source_folder):
                        full_path = os.path.join(self.source_folder, d)
                        if os.path.isdir(full_path):
                            folders_to_process.append({'name': d, 'path': full_path})
                else:
                    folders_to_process.append({'name': os.path.basename(self.source_folder), 'path': self.source_folder})
                    base_path = os.path.dirname(self.source_folder)

            for folder_info in folders_to_process:
                folder_name, full_path = folder_info['name'], folder_info['path']
                total_size = sum(os.path.getsize(os.path.join(dp, f)) for dp, dn, fn in os.walk(full_path) for f in fn)
                size_str = f"{total_size / 1024 / 1024:.2f} MB" if total_size > 0 else "0 MB"

                zip_filename = f"{folder_name}.zip"
                zip_path = os.path.join(self.output_folder, zip_filename)
                status = "⚠️ Exists" if os.path.exists(zip_path) else "✅ Ready"
                
                self.zip_tasks.append({'name': folder_name, 'path': full_path, 'zip_path': zip_path, 'status': status})
                self.preview_tree.insert('', 'end', values=(folder_name, os.path.relpath(full_path, base_path), size_str, status))

            if any(task['status'] == "✅ Ready" for task in self.zip_tasks):
                self.zip_btn.config(state='normal')
            self.status_var.set(f"Scan complete. Found {len(self.zip_tasks)} folder(s) to zip.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan folders: {e}")
            self.status_var.set("Error during folder scan.")

    def search_folders(self, *args):
        query = self.search_var.get().lower()
        self.preview_tree.delete(*self.preview_tree.get_children())

        if not self.zip_tasks:
            return

        base_path = self.source_folder
        if self.zip_top_level_only.get():
            base_path = self.source_folder
        elif self.source_folder:
            base_path = os.path.dirname(self.source_folder)

        def get_size_str(path):
            try:
                total_size = sum(os.path.getsize(os.path.join(dp, f)) for dp, dn, fn in os.walk(path) for f in fn)
                return f"{total_size / 1024 / 1024:.2f} MB" if total_size > 0 else "0 MB"
            except Exception:
                return "N/A"

        filtered_tasks = self.zip_tasks
        if query:
            filtered_tasks = [task for task in self.zip_tasks if query in task['name'].lower()]

        for task in filtered_tasks:
            size_str = get_size_str(task['path'])
            rel_path = os.path.relpath(task['path'], base_path) if base_path else task['path']
            self.preview_tree.insert('', 'end', values=(task['name'], rel_path, size_str, task['status']))

    def create_zip_files(self):
        tasks_to_run = [task for task in self.zip_tasks if task['status'] == "✅ Ready"]
        if not tasks_to_run:
            messagebox.showinfo("Info", "No new zip archives to create.")
            return
        if not messagebox.askyesno("Confirm Zipping", f"Are you sure you want to create {len(tasks_to_run)} zip archive(s)?"):
            return
        
        self.zip_btn.config(state='disabled')
        self.progress_bar['maximum'] = len(tasks_to_run)
        self.progress_bar['value'] = 0
        thread = threading.Thread(target=self._create_zip_files_thread, args=(tasks_to_run,))
        thread.start()

    def _create_zip_files_thread(self, tasks):
        processed_count = 0
        errors = []
        for i, task in enumerate(tasks):
            try:
                self.parent.after(0, self.status_var.set, f"Zipping {task['name']}...")
                with zipfile.ZipFile(task['zip_path'], 'w', zipfile.ZIP_DEFLATED) as zipf:
                    base_folder = os.path.dirname(task['path']) if self.include_root_dir.get() else task['path']
                    for root, _, files in os.walk(task['path']):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, base_folder)
                            zipf.write(file_path, arcname)
                task['status'] = "✅ Zipped"
                processed_count += 1
            except Exception as e:
                task['status'] = "❌ Error"
                errors.append(f"{task['name']}: {e}")
            finally:
                self.parent.after(0, self.progress_bar.config, {'value': i + 1})
        
        self.parent.after(0, self.finalize_zipping, processed_count, errors)

    def finalize_zipping(self, processed_count, errors):
        self.scan_folders() # Refresh the preview
        if errors:
            messagebox.showwarning("Zipping Complete with Errors", f"Zipped {processed_count} folders.\n\nErrors:\n" + "\n".join(errors))
        else:
            messagebox.showinfo("Success", f"Successfully created {processed_count} zip archive(s)!")
        self.status_var.set(f"Zipping complete. {processed_count} archives created.")
        self.zip_btn.config(state='disabled')
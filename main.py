import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
from modules.worksheet_sync import WorksheetSyncModule
from modules.bulk_rename import BulkRenameModule
from modules.multi_zip import MultiZipModule
from modules.folder_creator import FolderCreatorModule

class ModernButton(tk.Button):
    """Custom modern button with hover effects"""
    def __init__(self, parent, text, command, width=180, height=50, bg="#4A90E2", hover_bg="#357ABD", **kwargs):
        super().__init__(parent, text=text, command=command, 
                        width=width//10, height=height//20,  # Convert pixels to character units
                        font=("Segoe UI", 11, "bold"),
                        fg="white", bg=bg,
                        relief='flat', bd=0,
                        cursor="hand2",
                        activebackground=hover_bg,
                        activeforeground="white",
                        **kwargs)
        
        # Store colors for hover effects
        self.bg = bg
        self.hover_bg = hover_bg
        
        # Bind hover effects
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        
    def on_enter(self, event):
        """Mouse enter event"""
        self.configure(bg=self.hover_bg)
        
    def on_leave(self, event):
        """Mouse leave event"""
        self.configure(bg=self.bg)

class EruStudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EruStudio")
        self.root.configure(bg='#1a1a1a')
        self.root.state('zoomed')
        self.root.minsize(1024, 768)
        
        # Set the application icon
        self.set_icon(self.root)
        
        self.setup_styles()
        self.current_module = None
        self.module_cards = []
        self.is_single_column = False  # State tracker
        self.create_widgets()

        # Bind resize event
        self.root.bind("<Configure>", self.on_resize)
        
    def set_icon(self, window):
        """Set the application icon with multiple fallback methods"""
        # Try multiple possible icon paths
        icon_paths = []
        
        # Get base paths to check
        base_paths = []
        
        # Add current working directory
        base_paths.append(os.path.abspath('.'))
        
        # Add script directory
        if getattr(sys, 'frozen', False):
            # Running in PyInstaller bundle
            base_paths.append(sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable))
        else:
            # Running in a normal Python environment
            base_paths.append(os.path.dirname(os.path.abspath(__file__)))
        
        # Add executable directory
        if hasattr(sys, 'executable'):
            base_paths.append(os.path.dirname(sys.executable))
        
        # Add current directory
        base_paths.append(os.getcwd())
        
        # Generate possible icon paths
        for base in set(base_paths):  # Remove duplicates
            icon_paths.extend([
                os.path.join(base, 'icon.ico'),
                os.path.join(base, 'assets', 'icon.ico'),
                os.path.join(base, '..', 'icon.ico'),
                os.path.join(base, '..', 'assets', 'icon.ico')
            ])
        
        # Add absolute path to assets/icon.ico
        icon_paths.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.ico'))
        
        # Remove duplicates and non-existent paths
        icon_paths = list(dict.fromkeys(icon_paths))  # Remove duplicates while preserving order
        
        for icon_path in icon_paths:
            try:
                if os.path.exists(icon_path):
                    window.iconbitmap(icon_path)
                    print(f"Icon loaded from: {icon_path}")
                    return True
            except Exception as e:
                print(f"Warning: Could not load icon from {icon_path}: {e}")
        
        print("Warning: Could not set application icon - icon file not found in any standard location")
        return False

    def setup_styles(self):
        """Configure custom styles for the application"""
        style = ttk.Style()
        style.theme_use('clam')

        # Configure custom styles
        style.configure('Title.TLabel', font=('Segoe UI', 32, 'bold'), foreground='#ffffff')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 16), foreground='#b0b0b0')
        style.configure('Module.TFrame', background='#2d2d2d', relief='flat')
        style.configure('Header.TFrame', background='#008080') # Teal color

    def create_widgets(self):
        """Create the main application interface"""
        main_container = tk.Frame(self.root, bg='#1a1a1a')
        main_container.pack(fill='both', expand=True)

        header_frame = tk.Frame(main_container, bg='#008080') # Teal color
        header_frame.pack(fill='x')

        header_content = tk.Frame(header_frame, bg='#008080') # Teal color
        header_content.pack(expand=True, fill='both', padx=20, pady=20)

        tk.Label(header_content, text="EruStudio", font=('Segoe UI', 36, 'bold'), fg='#ffffff', bg='#008080').pack(pady=(0, 10))
        tk.Label(header_content, text="Professional File Management Suite", font=('Segoe UI', 14), fg='#e8f4fd', bg='#008080').pack()

        content_frame = tk.Frame(main_container, bg='#1a1a1a')
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)

        welcome_frame = tk.Frame(content_frame, bg='#1a1a1a')
        welcome_frame.pack(fill='x', pady=(0, 30))

        tk.Label(welcome_frame, text="Choose a module to get started:", font=('Segoe UI', 14), fg='#e0e0e0', bg='#1a1a1a').pack()

        self.modules_frame = tk.Frame(content_frame, bg='#1a1a1a')
        self.modules_frame.pack(fill='both', expand=True, pady=(20, 0))

        self.modules_frame.grid_columnconfigure(0, weight=1)
        self.modules_frame.grid_columnconfigure(1, weight=1)
        self.modules_frame.grid_rowconfigure(0, weight=1)
        self.modules_frame.grid_rowconfigure(1, weight=1)

        module_definitions = [
            ("📊", "Worksheet Sync", "Sync worksheets based on a common header", "#FF6B6B", self.open_worksheet_sync),
            ("🔄", "Bulk Rename", "Rename files using an Excel template", "#20c997", self.open_bulk_rename),
            ("📦", "Multi-Zip", "Create multiple zip archives from folders", "#20c997", self.open_multi_zip),
            ("📁", "Folder Creator", "Create folder structures from a template", "#96CEB4", self.open_folder_creator)
        ]

        positions = [(0, 0), (0, 1), (1, 0), (1, 1)]
        for i, (icon, title, desc, color, cmd) in enumerate(module_definitions):
            card = self.create_module_card(self.modules_frame, positions[i][0], positions[i][1], icon, title, desc, color, cmd)
            self.module_cards.append(card)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self.root, textvariable=self.status_var, relief='flat', anchor='w', bg='#2d2d2d', fg='#b0b0b0', font=('Segoe UI', 10)).pack(side='bottom', fill='x')

    def create_module_card(self, parent, row, col, icon, title, description, color, command):
        card_frame = tk.Frame(parent, bg='#2d2d2d', relief='flat', bd=0, cursor='hand2')
        card_frame.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')

        content_frame = tk.Frame(card_frame, bg='#2d2d2d', cursor='hand2')
        content_frame.pack(expand=True, fill='both', padx=20, pady=20)

        header_frame = tk.Frame(content_frame, bg='#2d2d2d', cursor='hand2')
        header_frame.pack(fill='x', pady=(0, 10))

        tk.Label(header_frame, text=icon, font=('Segoe UI', 28), fg=color, bg='#2d2d2d', cursor='hand2').pack(side='left', padx=(0, 15))
        tk.Label(header_frame, text=title, font=('Segoe UI', 16, 'bold'), fg='#ffffff', bg='#2d2d2d', cursor='hand2').pack(side='left')
        tk.Label(content_frame, text=description, font=('Segoe UI', 10), fg='#b0b0b0', bg='#2d2d2d', wraplength=200, justify='left', cursor='hand2').pack(anchor='w', pady=(0, 15))
        tk.Label(content_frame, text="Click to open", font=('Segoe UI', 8), fg='#888888', bg='#2d2d2d', cursor='hand2').pack(anchor='w')

        def on_enter(e): card_frame.config(bg='#3d3d3d'); content_frame.config(bg='#3d3d3d'); [w.config(bg='#3d3d3d') for w in content_frame.winfo_children()]
        def on_leave(e): card_frame.config(bg='#2d2d2d'); content_frame.config(bg='#2d2d2d'); [w.config(bg='#2d2d2d') for w in content_frame.winfo_children()]
        
        for widget in [card_frame, content_frame] + content_frame.winfo_children() + header_frame.winfo_children():
            widget.bind('<Enter>', on_enter)
            widget.bind('<Leave>', on_leave)
            widget.bind('<Button-1>', lambda e: command())

        return card_frame

    def open_module(self, module_class, title):
        if self.current_module:
            self.current_module.destroy()

        module_window = tk.Toplevel(self.root)
        module_window.title(f"EruStudio - {title}")
        
        # Set the window icon using our robust method
        self.set_icon(module_window)
            
        module_window.state('zoomed')
        module_window.configure(bg='#1a1a1a')
        module_window.transient(self.root)
        module_window.grab_set()
        module_window.minsize(800, 600)
        module_window.resizable(True, True)

        # Set the window icon again after window is fully created
        module_window.after(100, lambda: self.set_icon(module_window))

        self.current_module = module_window
        self.status_var.set(f"Module opened: {title}")
        module_class(module_window)

        module_window.protocol("WM_DELETE_WINDOW", lambda: (self.status_var.set("Ready"), setattr(self, 'current_module', None), module_window.destroy()))

    def open_worksheet_sync(self): self.open_module(WorksheetSyncModule, "Worksheet Sync")
    def open_bulk_rename(self): self.open_module(BulkRenameModule, "Bulk Rename")
    def open_multi_zip(self): self.open_module(MultiZipModule, "Multi-Zip")
    def open_folder_creator(self): self.open_module(FolderCreatorModule, "Folder Creator")

    def on_resize(self, event):
        width = self.root.winfo_width()
        should_be_single_column = width < 800

        if should_be_single_column and not self.is_single_column:
            self.modules_frame.grid_columnconfigure(1, weight=0)
            for i, card in enumerate(self.module_cards):
                card.grid_forget(); card.grid(row=i, column=0, padx=10, pady=10, sticky='nsew')
            self.is_single_column = True
        elif not should_be_single_column and self.is_single_column:
            self.modules_frame.grid_columnconfigure(1, weight=1)
            positions = [(0, 0), (0, 1), (1, 0), (1, 1)]
            for i, card in enumerate(self.module_cards):
                card.grid_forget(); card.grid(row=positions[i][0], column=positions[i][1], padx=10, pady=10, sticky='nsew')
            self.is_single_column = False

def main():
    root = tk.Tk()
    app = EruStudioApp(root)
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    root.resizable(True, True)
    root.mainloop()

if __name__ == "__main__":
    main()
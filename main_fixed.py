import tkinter as tk
from tkinter import ttk, messagebox
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
        self.root.title("EruStudio - Professional File Management Suite")
        self.root.geometry("1200x800")
        self.root.configure(bg='#1a1a1a')
        
        # Set application icon (if available)
        try:
            self.root.iconbitmap("assets/icon.ico")
        except:
            pass
        
        self.setup_styles()
        self.create_widgets()
        self.current_module = None
        
    def setup_styles(self):
        """Configure custom styles for the application"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure custom styles
        style.configure('Title.TLabel', font=('Segoe UI', 32, 'bold'), foreground='#ffffff')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 16), foreground='#b0b0b0')
        style.configure('Module.TFrame', background='#2d2d2d', relief='flat')
        style.configure('Header.TFrame', background='#4A90E2')
        
    def create_widgets(self):
        """Create the main application interface"""
        # Main container with gradient background
        main_container = tk.Frame(self.root, bg='#1a1a1a')
        main_container.pack(fill='both', expand=True, padx=0, pady=0)
        
        # Header with gradient effect
        header_frame = tk.Frame(main_container, bg='#4A90E2', height=200)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header_frame, bg='#4A90E2')
        header_content.pack(expand=True, fill='both', padx=40, pady=30)
        
        # Title with shadow effect
        title_label = tk.Label(header_content, text="EruStudio", 
                              font=('Segoe UI', 48, 'bold'), 
                              fg='#ffffff', bg='#4A90E2')
        title_label.pack(pady=(0, 10))
        
        subtitle_label = tk.Label(header_content, text="Professional File Management Suite", 
                                 font=('Segoe UI', 18), 
                                 fg='#e8f4fd', bg='#4A90E2')
        subtitle_label.pack()
        
        # Main content area
        content_frame = tk.Frame(main_container, bg='#1a1a1a')
        content_frame.pack(fill='both', expand=True, padx=40, pady=40)
        
        # Welcome message
        welcome_frame = tk.Frame(content_frame, bg='#1a1a1a')
        welcome_frame.pack(fill='x', pady=(0, 40))
        
        welcome_label = tk.Label(welcome_frame, 
                                text="Choose a module to get started with your file management tasks:",
                                font=('Segoe UI', 16), 
                                fg='#e0e0e0', bg='#1a1a1a')
        welcome_label.pack()
        
        # Module grid with better spacing
        modules_frame = tk.Frame(content_frame, bg='#1a1a1a')
        modules_frame.pack(fill='both', expand=True, pady=(20, 0))
        
        # Configure grid weights
        modules_frame.grid_columnconfigure(0, weight=1)
        modules_frame.grid_columnconfigure(1, weight=1)
        modules_frame.grid_rowconfigure(0, weight=1)
        modules_frame.grid_rowconfigure(1, weight=1)
        
        # Module 1: Worksheet Sync
        self.create_module_card(modules_frame, 0, 0, 
                               "📊", "Worksheet Sync", 
                               "Sync all worksheets of a workbook\nbased on selected header row",
                               "#FF6B6B", "#FF5252", self.open_worksheet_sync)
        
        # Module 2: Bulk Rename
        self.create_module_card(modules_frame, 0, 1,
                               "🔄", "Bulk Rename", 
                               "Rename files/folders based on\nuploaded Excel template",
                               "#4ECDC4", "#26A69A", self.open_bulk_rename)
        
        # Module 3: Multi-Zip
        self.create_module_card(modules_frame, 1, 0,
                               "📦", "Multi-Zip", 
                               "Create multiple zip files\nper folder",
                               "#45B7D1", "#2196F3", self.open_multi_zip)
        
        # Module 4: Folder Creator
        self.create_module_card(modules_frame, 1, 1,
                               "📁", "Folder Creator", 
                               "Create multiple folders based on\nuploaded Excel template",
                               "#96CEB4", "#4CAF50", self.open_folder_creator)
        
        # Templates section
        templates_frame = tk.Frame(content_frame, bg='#2d2d2d', relief='flat', bd=0)
        templates_frame.pack(fill='x', pady=(40, 0))
        
        # Templates header
        templates_header = tk.Frame(templates_frame, bg='#2d2d2d')
        templates_header.pack(fill='x', padx=20, pady=(20, 10))
        
        tk.Label(templates_header, text="📋 Built-in Excel Templates", 
                font=('Segoe UI', 18, 'bold'), 
                fg='#ffffff', bg='#2d2d2d').pack(side='left')
        
        # Templates info
        templates_info = tk.Frame(templates_frame, bg='#2d2d2d')
        templates_info.pack(fill='x', padx=20, pady=(0, 20))
        
        templates_text = """Ready-to-use Excel templates are included with EruStudio:
• bulk_rename_template.xlsx - For file renaming operations
• folder_creator_template.xlsx - For creating folder structures  
• worksheet_sync_template.xlsx - For worksheet synchronization
• multi_zip_template.xlsx - For zip configuration"""
        
        tk.Label(templates_info, text=templates_text, 
                font=('Segoe UI', 11), 
                fg='#b0b0b0', bg='#2d2d2d', 
                justify='left').pack(anchor='w')
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                             relief='flat', anchor='w', 
                             bg='#2d2d2d', fg='#b0b0b0',
                             font=('Segoe UI', 10))
        status_bar.pack(side='bottom', fill='x', padx=0, pady=0)
        
    def create_module_card(self, parent, row, col, icon, title, description, color, hover_color, command):
        """Create a modern module card that's entirely clickable"""
        # Card container - make it entirely clickable
        card_frame = tk.Frame(parent, bg='#2d2d2d', relief='flat', bd=0, cursor='hand2')
        card_frame.grid(row=row, column=col, padx=15, pady=15, sticky='nsew')
        
        # Card content
        content_frame = tk.Frame(card_frame, bg='#2d2d2d', cursor='hand2')
        content_frame.pack(expand=True, fill='both', padx=25, pady=25)
        
        # Icon and title
        header_frame = tk.Frame(content_frame, bg='#2d2d2d', cursor='hand2')
        header_frame.pack(fill='x', pady=(0, 15))
        
        icon_label = tk.Label(header_frame, text=icon, 
                             font=('Segoe UI', 36), 
                             fg=color, bg='#2d2d2d', cursor='hand2')
        icon_label.pack(side='left', padx=(0, 15))
        
        title_label = tk.Label(header_frame, text=title, 
                               font=('Segoe UI', 18, 'bold'), 
                               fg='#ffffff', bg='#2d2d2d', cursor='hand2')
        title_label.pack(side='left')
        
        # Description
        desc_label = tk.Label(content_frame, text=description, 
                             font=('Segoe UI', 11), 
                             fg='#b0b0b0', bg='#2d2d2d', 
                             wraplength=220, justify='center',
                             padx=10, cursor='hand2')
        desc_label.pack(pady=(0, 20))
        
        # Click instruction
        click_label = tk.Label(content_frame, text="Click anywhere on this card to open", 
                              font=('Segoe UI', 9), 
                              fg='#888888', bg='#2d2d2d', cursor='hand2')
        click_label.pack()
        
        # Hover effect for entire card
        def on_enter(e):
            card_frame.configure(bg='#3d3d3d')
            content_frame.configure(bg='#3d3d3d')
            header_frame.configure(bg='#3d3d3d')
            desc_label.configure(bg='#3d3d3d')
            click_label.configure(bg='#3d3d3d')
            
        def on_leave(e):
            card_frame.configure(bg='#2d2d2d')
            content_frame.configure(bg='#2d2d2d')
            header_frame.configure(bg='#2d2d2d')
            desc_label.configure(bg='#2d2d2d')
            click_label.configure(bg='#2d2d2d')
        
        # Make entire card clickable
        def on_click(e):
            command()
            
        # Bind click events to all elements
        for widget in [card_frame, content_frame, header_frame, icon_label, title_label, desc_label, click_label]:
            widget.bind('<Enter>', on_enter)
            widget.bind('<Leave>', on_leave)
            widget.bind('<Button-1>', on_click)
        
    def open_worksheet_sync(self):
        """Open the Worksheet Sync module"""
        print("Opening Worksheet Sync module...")
        self.open_module(WorksheetSyncModule, "Worksheet Sync")
        
    def open_bulk_rename(self):
        """Open the Bulk Rename module"""
        print("Opening Bulk Rename module...")
        self.open_module(BulkRenameModule, "Bulk Rename")
        
    def open_multi_zip(self):
        """Open the Multi-Zip module"""
        print("Opening Multi-Zip module...")
        self.open_module(MultiZipModule, "Multi-Zip")
        
    def open_folder_creator(self):
        """Open the Folder Creator module"""
        print("Opening Folder Creator module...")
        self.open_module(FolderCreatorModule, "Folder Creator")
        
    def open_module(self, module_class, title):
        """Open a module in a new window"""
        try:
            print(f"Attempting to open module: {title}")
            
            if self.current_module:
                print("Closing existing module...")
                self.current_module.destroy()
                
            print("Creating new module window...")
            module_window = tk.Toplevel(self.root)
            module_window.title(f"EruStudio - {title}")
            module_window.geometry("1000x700")
            module_window.configure(bg='#1a1a1a')
            
            # Center the window
            module_window.transient(self.root)
            module_window.grab_set()
            
            # Make window resizable
            module_window.resizable(True, True)
            
            self.current_module = module_window
            self.status_var.set(f"Module opened: {title}")
            
            print("Creating module instance...")
            # Create and initialize the module
            module = module_class(module_window)
            print(f"Module {title} opened successfully!")
            
            # Handle window close
            def on_closing():
                print(f"Closing module: {title}")
                self.current_module = None
                self.status_var.set("Ready")
                module_window.destroy()
                
            module_window.protocol("WM_DELETE_WINDOW", on_closing)
            
        except Exception as e:
            print(f"Error opening module {title}: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open {title} module: {str(e)}")

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = EruStudioApp(root)
    
    # Center the main window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Make window resizable
    root.resizable(True, True)
    
    root.mainloop()

if __name__ == "__main__":
    main() 
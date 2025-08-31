import os
import sys
import tkinter as tk
from tkinter import messagebox

def verify_icon():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    icon_path = os.path.join('assets', 'icon.ico')
    
    # Check if icon file exists
    if not os.path.exists(icon_path):
        messagebox.showerror("Error", f"Icon file not found at: {icon_path}")
        return False
    
    # Try to set the icon
    try:
        root.iconbitmap(icon_path)
        messagebox.showinfo("Success", "Icon is valid and can be loaded!")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load icon: {str(e)}")
        return False

if __name__ == "__main__":
    verify_icon()

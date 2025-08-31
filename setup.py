#!/usr/bin/env python3
"""
EruStudio Setup Script
This script helps install dependencies and configure the application.
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"  {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"    ✓ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"    ✗ {description} failed: {e}")
        if e.stdout:
            print(f"      Output: {e.stdout}")
        if e.stderr:
            print(f"      Error: {e.stderr}")
        return False

def check_python_version():
    """Check if Python version is compatible"""
    print("Checking Python version...")
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"  ✗ Python {version.major}.{version.minor} is not supported")
        print("    Please install Python 3.8 or higher")
        return False
    else:
        print(f"  ✓ Python {version.major}.{version.minor}.{version.micro} is compatible")
        return True

def install_dependencies():
    """Install required Python packages"""
    print("\nInstalling dependencies...")
    
    # Upgrade pip first
    if not run_command("python -m pip install --upgrade pip", "Upgrading pip"):
        print("  Warning: Failed to upgrade pip, continuing...")
    
    # Install requirements
    if not run_command("pip install -r requirements.txt", "Installing packages from requirements.txt"):
        return False
    
    return True

def generate_templates():
    """Generate Excel templates for users"""
    print("\nGenerating enhanced Excel templates...")
    
    try:
        # Import template manager
        from templates.template_manager import TemplateManager
        manager = TemplateManager()
        manager.generate_all_templates()
        print("  ✓ Enhanced Excel templates generated successfully")
        return True
    except ImportError:
        print("  ✗ Could not import template manager")
        return False
    except Exception as e:
        print(f"  ✗ Failed to generate templates: {e}")
        return False

def create_sample_templates():
    """Create sample Excel templates for users (fallback method)"""
    print("\nCreating sample templates (fallback method)...")
    
    try:
        import pandas as pd
        
        # Create templates directory
        templates_dir = Path('templates')
        templates_dir.mkdir(exist_ok=True)
        
        # Create sample bulk rename template
        rename_data = {
            'Current Name': ['file1', 'file2', 'document1', 'image1'],
            'New Name': ['renamed_file1', 'renamed_file2', 'renamed_document1', 'renamed_image1']
        }
        rename_df = pd.DataFrame(rename_data)
        rename_path = templates_dir / 'bulk_rename_template.xlsx'
        rename_df.to_excel(rename_path, index=False)
        print(f"  ✓ Created {rename_path}")
        
        # Create sample folder creator template
        folder_data = {
            'Folder Name': ['Projects', 'Documents', 'Images', 'Backups'],
            'Parent Folder': ['Work', 'Work', 'Media', '']
        }
        folder_df = pd.DataFrame(folder_data)
        folder_path = templates_dir / 'folder_creator_template.xlsx'
        folder_df.to_excel(folder_path, index=False)
        print(f"  ✓ Created {folder_path}")
        
        return True
        
    except ImportError:
        print("  ✗ Could not create sample templates (pandas not available)")
        return False
    except Exception as e:
        print(f"  ✗ Failed to create templates: {e}")
        return False

def main():
    """Main setup function"""
    print("=" * 60)
    print("EruStudio Setup Script")
    print("=" * 60)
    
    # Check Python version
    if not check_python_version():
        print("\nSetup failed: Incompatible Python version")
        return 1
    
    # Install dependencies
    if not install_dependencies():
        print("\nSetup failed: Could not install dependencies")
        return 1
    
    # Generate templates
    if not generate_templates():
        print("  Warning: Could not generate templates, trying fallback method...")
        if not create_sample_templates():
            print("  Warning: Could not create templates")
    
    print("\n" + "=" * 60)
    print("Setup completed successfully!")
    print("=" * 60)
    print("\nTo launch EruStudio:")
    print("  • Run: python main.py")
    print("  • Or double-click: run_erustudio.bat")
    print("  • Or run: powershell -ExecutionPolicy Bypass -File run_erustudio.ps1")
    print("\nExcel templates have been created:")
    print("  • bulk_rename_template.xlsx")
    print("  • folder_creator_template.xlsx")
    print("  • worksheet_sync_template.xlsx")
    print("  • multi_zip_template.xlsx")
    print("\nFor more information, see README.md")
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 
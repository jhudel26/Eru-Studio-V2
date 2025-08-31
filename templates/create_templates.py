#!/usr/bin/env python3
"""
Template Generator for EruStudio
Creates sample Excel templates for all modules
"""

import pandas as pd
import os
from pathlib import Path

def create_bulk_rename_template():
    """Create sample bulk rename template"""
    data = {
        'Current Name': [
            'document_old',
            'image_old', 
            'file_old',
            'report_old',
            'data_old',
            'backup_old',
            'archive_old',
            'temp_old'
        ],
        'New Name': [
            'document_new',
            'image_new',
            'file_new', 
            'report_new',
            'data_new',
            'backup_new',
            'archive_new',
            'temp_new'
        ],
        'Category': [
            'Documents',
            'Images',
            'Files',
            'Reports',
            'Data',
            'Backups',
            'Archives',
            'Temporary'
        ]
    }
    
    df = pd.DataFrame(data)
    return df

def create_folder_creator_template():
    """Create sample folder creator template"""
    data = {
        'Folder Name': [
            'Projects',
            'Documents',
            'Images',
            'Videos',
            'Music',
            'Downloads',
            'Backups',
            'Archives'
        ],
        'Parent Folder': [
            'Work',
            'Work',
            'Media',
            'Media',
            'Media',
            'User',
            'System',
            'System'
        ],
        'Description': [
            'Work projects and assignments',
            'Important documents and files',
            'Photos and graphics',
            'Video files and recordings',
            'Audio files and music',
            'Downloaded files',
            'System backups',
            'Long-term archives'
        ]
    }
    
    df = pd.DataFrame(data)
    return df

def create_worksheet_sync_template():
    """Create sample worksheet sync template"""
    # Create multiple sheets with different data structures
    sheets = {}
    
    # Sheet 1: Sales Data
    sales_data = {
        'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones'],
        'Price': [999.99, 29.99, 79.99, 299.99, 149.99],
        'Quantity': [10, 50, 25, 15, 30],
        'Category': ['Electronics', 'Accessories', 'Accessories', 'Electronics', 'Accessories']
    }
    sheets['Sales'] = pd.DataFrame(sales_data)
    
    # Sheet 2: Customer Data
    customer_data = {
        'Customer ID': ['C001', 'C002', 'C003', 'C004', 'C005'],
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson'],
        'Email': ['john@email.com', 'jane@email.com', 'bob@email.com', 'alice@email.com', 'charlie@email.com'],
        'Phone': ['555-0101', '555-0102', '555-0103', '555-0104', '555-0105']
    }
    sheets['Customers'] = pd.DataFrame(customer_data)
    
    # Sheet 3: Inventory
    inventory_data = {
        'Item Code': ['I001', 'I002', 'I003', 'I004', 'I005'],
        'Item Name': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones'],
        'Stock': [25, 100, 50, 30, 60],
        'Location': ['Warehouse A', 'Warehouse B', 'Warehouse A', 'Warehouse C', 'Warehouse B']
    }
    sheets['Inventory'] = pd.DataFrame(inventory_data)
    
    return sheets

def create_multi_zip_template():
    """Create sample multi-zip template"""
    data = {
        'Folder Name': [
            'Project_A',
            'Project_B', 
            'Project_C',
            'Documents_2024',
            'Images_Q1',
            'Backups_Jan',
            'Archives_2023',
            'Temp_Files'
        ],
        'Compression Level': [6, 8, 4, 9, 5, 7, 9, 1],
        'Include Subfolders': [True, True, False, True, True, True, True, False],
        'Description': [
            'Main project files with subfolders',
            'Secondary project with full structure',
            'Simple project without subfolders',
            'Important documents with maximum compression',
            'Image files with medium compression',
            'Backup files with high compression',
            'Archive files with maximum compression',
            'Temporary files with minimal compression'
        ]
    }
    
    df = pd.DataFrame(data)
    return df

def generate_all_templates():
    """Generate all template files"""
    # Create templates directory
    templates_dir = Path('templates')
    templates_dir.mkdir(exist_ok=True)
    
    print("Generating Excel templates...")
    
    # Bulk Rename Template
    try:
        rename_df = create_bulk_rename_template()
        rename_path = templates_dir / 'bulk_rename_template.xlsx'
        rename_df.to_excel(rename_path, index=False, sheet_name='Rename_Mapping')
        print(f"  ✓ Created {rename_path}")
    except Exception as e:
        print(f"  ✗ Failed to create bulk rename template: {e}")
    
    # Folder Creator Template
    try:
        folder_df = create_folder_creator_template()
        folder_path = templates_dir / 'folder_creator_template.xlsx'
        folder_df.to_excel(folder_path, index=False, sheet_name='Folder_Structure')
        print(f"  ✓ Created {folder_path}")
    except Exception as e:
        print(f"  ✗ Failed to create folder creator template: {e}")
    
    # Worksheet Sync Template
    try:
        sync_sheets = create_worksheet_sync_template()
        sync_path = templates_dir / 'worksheet_sync_template.xlsx'
        with pd.ExcelWriter(sync_path, engine='openpyxl') as writer:
            for sheet_name, df in sync_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"  ✓ Created {sync_path}")
    except Exception as e:
        print(f"  ✗ Failed to create worksheet sync template: {e}")
    
    # Multi-Zip Template
    try:
        zip_df = create_multi_zip_template()
        zip_path = templates_dir / 'multi_zip_template.xlsx'
        zip_df.to_excel(zip_path, index=False, sheet_name='Zip_Configuration')
        print(f"  ✓ Created {zip_path}")
    except Exception as e:
        print(f"  ✗ Failed to create multi-zip template: {e}")
    
    print("\nAll templates generated successfully!")
    print("Templates are located in the 'templates' folder.")

if __name__ == "__main__":
    generate_all_templates() 
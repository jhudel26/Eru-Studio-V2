#!/usr/bin/env python3
"""
Template Manager for EruStudio
Generates and manages templates for each module
"""

import pandas as pd
import os
from pathlib import Path

class TemplateManager:
    def __init__(self):
        self.templates_dir = Path('templates')
        self.templates_dir.mkdir(exist_ok=True)
        
    def create_worksheet_sync_template(self):
        """Create worksheet sync template with sample data"""
        # Create multiple sheets with different data structures
        sheets = {}
        
        # Sheet 1: Sales Data
        sales_data = {
            'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones', 'Webcam', 'Speaker', 'Microphone'],
            'Price': [999.99, 29.99, 79.99, 299.99, 149.99, 89.99, 199.99, 129.99],
            'Quantity': [10, 50, 25, 15, 30, 20, 12, 18],
            'Category': ['Electronics', 'Accessories', 'Accessories', 'Electronics', 'Accessories', 'Accessories', 'Audio', 'Audio'],
            'Supplier': ['TechCorp', 'AccessPro', 'AccessPro', 'TechCorp', 'AudioMax', 'AccessPro', 'AudioMax', 'AudioMax']
        }
        sheets['Sales'] = pd.DataFrame(sales_data)
        
        # Sheet 2: Customer Data
        customer_data = {
            'Customer ID': ['C001', 'C002', 'C003', 'C004', 'C005', 'C006', 'C007', 'C008'],
            'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson', 'Diana Prince', 'Eve Adams', 'Frank Miller'],
            'Email': ['john@email.com', 'jane@email.com', 'bob@email.com', 'alice@email.com', 'charlie@email.com', 'diana@email.com', 'eve@email.com', 'frank@email.com'],
            'Phone': ['555-0101', '555-0102', '555-0103', '555-0104', '555-0105', '555-0106', '555-0107', '555-0108'],
            'Address': ['123 Main St', '456 Oak Ave', '789 Pine Rd', '321 Elm St', '654 Maple Dr', '987 Cedar Ln', '147 Birch Way', '258 Spruce Ct']
        }
        sheets['Customers'] = pd.DataFrame(customer_data)
        
        # Sheet 3: Inventory
        inventory_data = {
            'Item Code': ['I001', 'I002', 'I003', 'I004', 'I005', 'I006', 'I007', 'I008'],
            'Item Name': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones', 'Webcam', 'Speaker', 'Microphone'],
            'Stock': [25, 100, 50, 30, 60, 40, 15, 25],
            'Location': ['Warehouse A', 'Warehouse B', 'Warehouse A', 'Warehouse C', 'Warehouse B', 'Warehouse A', 'Warehouse C', 'Warehouse B'],
            'Last Updated': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19', '2024-01-20', '2024-01-21', '2024-01-22']
        }
        sheets['Inventory'] = pd.DataFrame(inventory_data)
        
        # Sheet 4: Orders
        orders_data = {
            'Order ID': ['O001', 'O002', 'O003', 'O004', 'O005', 'O006', 'O007', 'O008'],
            'Customer ID': ['C001', 'C002', 'C003', 'C004', 'C005', 'C006', 'C007', 'C008'],
            'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones', 'Webcam', 'Speaker', 'Microphone'],
            'Quantity': [1, 2, 1, 1, 1, 1, 1, 1],
            'Order Date': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19', '2024-01-20', '2024-01-21', '2024-01-22'],
            'Status': ['Shipped', 'Processing', 'Delivered', 'Shipped', 'Processing', 'Delivered', 'Shipped', 'Processing']
        }
        sheets['Orders'] = pd.DataFrame(orders_data)
        
        return sheets
    
    def create_bulk_rename_template(self):
        """Create bulk rename template with comprehensive examples"""
        data = {
            'Current Name': [
                'document_old_v1',
                'image_old_2023', 
                'file_old_backup',
                'report_old_final',
                'data_old_raw',
                'backup_old_jan',
                'archive_old_2023',
                'temp_old_draft',
                'presentation_old',
                'spreadsheet_old'
            ],
            'New Name': [
                'document_new_v2',
                'image_new_2024',
                'file_new_current',
                'report_new_final',
                'data_new_processed',
                'backup_new_feb',
                'archive_new_2024',
                'temp_new_final',
                'presentation_new',
                'spreadsheet_new'
            ],
            'Category': [
                'Documents',
                'Images',
                'Files',
                'Reports',
                'Data',
                'Backups',
                'Archives',
                'Temporary',
                'Presentations',
                'Spreadsheets'
            ],
            'Priority': [
                'High',
                'Medium',
                'Low',
                'High',
                'Medium',
                'Low',
                'Medium',
                'Low',
                'High',
                'Medium'
            ]
        }
        
        return pd.DataFrame(data)
    
    def create_folder_creator_template(self):
        """Create folder creator template with organizational structure"""
        data = {
            'Folder Name': [
                'Projects',
                'Documents',
                'Images',
                'Videos',
                'Music',
                'Downloads',
                'Backups',
                'Archives',
                'Templates',
                'Reports',
                'Data',
                'Media'
            ],
            'Parent Folder': [
                'Work',
                'Work',
                'Media',
                'Media',
                'Media',
                'User',
                'System',
                'System',
                'Work',
                'Work',
                'Work',
                'User'
            ],
            'Description': [
                'Work projects and assignments',
                'Important documents and files',
                'Photos and graphics',
                'Video files and recordings',
                'Audio files and music',
                'Downloaded files',
                'System backups',
                'Long-term archives',
                'Reusable templates',
                'Business reports',
                'Data analysis files',
                'User media content'
            ],
            'Access Level': [
                'Private',
                'Private',
                'Shared',
                'Shared',
                'Shared',
                'Public',
                'System',
                'System',
                'Shared',
                'Private',
                'Private',
                'Public'
            ]
        }
        
        return pd.DataFrame(data)
    
    def create_multi_zip_template(self):
        """Create multi-zip template with configuration options"""
        data = {
            'Folder Name': [
                'Project_A',
                'Project_B', 
                'Project_C',
                'Documents_2024',
                'Images_Q1',
                'Backups_Jan',
                'Archives_2023',
                'Temp_Files',
                'Client_Data',
                'Internal_Reports'
            ],
            'Compression Level': [6, 8, 4, 9, 5, 7, 9, 1, 6, 8],
            'Include Subfolders': [True, True, False, True, True, True, True, False, True, False],
            'Description': [
                'Main project files with subfolders',
                'Secondary project with full structure',
                'Simple project without subfolders',
                'Important documents with maximum compression',
                'Image files with medium compression',
                'Backup files with high compression',
                'Archive files with maximum compression',
                'Temporary files with minimal compression',
                'Client data with standard compression',
                'Internal reports with high compression'
            ],
            'Password Protected': [False, False, False, True, False, True, True, False, True, False]
        }
        
        return pd.DataFrame(data)
    
    def generate_all_templates(self):
        """Generate all template files with improved content"""
        print("Generating enhanced Excel templates...")
        
        # Worksheet Sync Template
        try:
            sync_sheets = self.create_worksheet_sync_template()
            sync_path = self.templates_dir / 'worksheet_sync_template.xlsx'
            with pd.ExcelWriter(sync_path, engine='openpyxl') as writer:
                for sheet_name, df in sync_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  ✓ Created {sync_path}")
        except Exception as e:
            print(f"  ✗ Failed to create worksheet sync template: {e}")
        
        # Bulk Rename Template
        try:
            rename_df = self.create_bulk_rename_template()
            rename_path = self.templates_dir / 'bulk_rename_template.xlsx'
            rename_df.to_excel(rename_path, index=False, sheet_name='Rename_Mapping')
            print(f"  ✓ Created {rename_path}")
        except Exception as e:
            print(f"  ✗ Failed to create bulk rename template: {e}")
        
        # Folder Creator Template
        try:
            folder_df = self.create_folder_creator_template()
            folder_path = self.templates_dir / 'folder_creator_template.xlsx'
            folder_df.to_excel(folder_path, index=False, sheet_name='Folder_Structure')
            print(f"  ✓ Created {folder_path}")
        except Exception as e:
            print(f"  ✗ Failed to create folder creator template: {e}")
        
        # Multi-Zip Template
        try:
            zip_df = self.create_multi_zip_template()
            zip_path = self.templates_dir / 'multi_zip_template.xlsx'
            zip_df.to_excel(zip_path, index=False, sheet_name='Zip_Configuration')
            print(f"  ✓ Created {zip_path}")
        except Exception as e:
            print(f"  ✗ Failed to create multi-zip template: {e}")
        
        print("\nAll enhanced templates generated successfully!")
        print("Templates are located in the 'templates' folder.")
        print("\nEach template now includes:")
        print("  • More comprehensive sample data")
        print("  • Additional useful columns")
        print("  • Better organization and structure")
        print("  • Real-world examples")

def main():
    """Generate all templates"""
    manager = TemplateManager()
    manager.generate_all_templates()

if __name__ == "__main__":
    main() 
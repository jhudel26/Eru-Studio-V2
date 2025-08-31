# EruStudio - Professional File Management Suite

A comprehensive Windows application for advanced file management, Excel processing, and bulk operations with a modern, professional interface.

## ✨ Features

### 🎨 Modern User Interface
- **Dark Theme Design** with professional color scheme
- **Responsive Layout** that adapts to different screen sizes
- **Interactive Cards** with hover effects and smooth animations
- **Professional Typography** using Segoe UI font family
- **Color-coded Modules** for easy identification and navigation

### 📊 Worksheet Sync Module
- **Synchronize multiple worksheets** from a single Excel workbook
- **Customizable header row selection** for data alignment
- **Preview data** before synchronization
- **Export synced data** to new Excel files
- **Multi-worksheet selection** for batch processing

### 🔄 Bulk Rename Module
- **Rename files and folders** based on Excel templates
- **Column mapping** for current and new names
- **File extension preservation** options
- **Backup creation** before renaming
- **Conflict detection** and resolution
- **Preview before execution** for safety

### 📦 Multi-Zip Module
- **Create multiple zip files** per folder
- **Recursive zipping** with subfolder support
- **Configurable compression levels** (0-9)
- **Folder structure analysis** before zipping
- **Progress tracking** with visual indicators
- **Overwrite protection** options

### 📁 Folder Creator Module
- **Create multiple folders** from Excel templates
- **Nested folder structure** support
- **Parent-child relationships** based on template data
- **README.txt generation** in each folder
- **Conflict resolution** for existing folders
- **Preview generation** before creation

### 📋 Built-in Excel Templates
- **Ready-to-use templates** included with the application
- **bulk_rename_template.xlsx** - For file renaming operations
- **folder_creator_template.xlsx** - For creating folder structures
- **worksheet_sync_template.xlsx** - For worksheet synchronization
- **multi_zip_template.xlsx** - For zip configuration
- **Professional formatting** with sample data and descriptions

## 🚀 Installation

### Prerequisites
- Windows 10/11
- Python 3.8 or higher
- pip package manager

### Quick Setup
1. **Clone or download** the project files
2. **Run the setup script**:
   ```bash
   python setup.py
   ```
3. **Launch the application**:
   ```bash
   python main.py
   ```

### Manual Setup
1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
2. **Generate templates**:
   ```bash
   python templates/create_templates.py
   ```
3. **Run the application**:
   ```bash
   python main.py
   ```

## 🎯 Usage

### Getting Started
1. **Launch EruStudio** from the main window
2. **Choose a module** from the modern card interface
3. **Follow the on-screen instructions** for each module
4. **Use the preview features** to verify operations before execution
5. **Access built-in templates** for immediate use

### Worksheet Sync
1. Select an Excel workbook file
2. Choose the header row number
3. Select worksheets to synchronize
4. Preview the synced data
5. Export to a new file

### Bulk Rename
1. **Use built-in template** or upload your own Excel file
2. Select the source folder containing files to rename
3. Map the template columns (Current Name → New Name)
4. Generate preview and review changes
5. Execute the rename operation

### Multi-Zip
1. Select source folder to scan
2. Choose output folder for zip files
3. Configure compression and structure options
4. Scan folders to analyze structure
5. Create zip files for each folder

### Folder Creator
1. **Use built-in template** or upload your own Excel file
2. Select output location
3. Configure folder structure options
4. Generate preview of folder structure
5. Create folders based on template

## 📋 Excel Templates

### Built-in Templates
All templates are automatically generated during setup and include:

#### Bulk Rename Template
- **Current Name**: Original file names (without extensions)
- **New Name**: Target file names (without extensions)
- **Category**: Optional categorization for organization

#### Folder Creator Template
- **Folder Name**: Names of folders to create
- **Parent Folder**: Parent directory names (for nested structure)
- **Description**: Purpose and description of each folder

#### Worksheet Sync Template
- **Multiple sheets** with different data structures
- **Sample data** for testing and understanding
- **Professional formatting** ready for customization

#### Multi-Zip Template
- **Folder configurations** for zip operations
- **Compression settings** and options
- **Structure preferences** for different use cases

### Template Customization
- **Modify existing templates** to match your needs
- **Add new columns** for additional functionality
- **Use as reference** for creating your own templates
- **Professional formatting** maintained across all templates

## ⚙️ Configuration Options

### General Settings
- **File extensions**: Include or exclude file extensions during operations
- **Backup creation**: Automatic backup before destructive operations
- **Overwrite protection**: Prevent accidental overwrites
- **Progress tracking**: Visual progress indicators for long operations

### Zip Settings
- **Compression level**: 0 (no compression) to 9 (maximum compression)
- **Folder structure**: Include or exclude subfolder hierarchy
- **Recursive zipping**: Create nested zip structures

### UI Settings
- **Dark theme**: Professional dark color scheme
- **Responsive design**: Adapts to different screen sizes
- **Hover effects**: Interactive feedback for better UX
- **Modern typography**: Clean, readable text presentation

## 🛡️ Safety Features

- **Preview before execution** for all operations
- **Automatic backups** before file modifications
- **Conflict detection** and resolution
- **Progress tracking** for long operations
- **Error handling** with detailed error messages
- **Undo protection** through backup creation
- **Template validation** before processing

## 💻 System Requirements

- **Operating System**: Windows 10/11 (64-bit)
- **Python**: 3.8 or higher
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Storage**: 100MB free space for application
- **Display**: 1024x768 minimum resolution, 1920x1080 recommended

## 🔧 Troubleshooting

### Common Issues

**"Module not found" error**
- Ensure all dependencies are installed: `pip install -r requirements.txt`

**Excel file loading errors**
- Verify the file is not open in Excel
- Check file format (.xlsx or .xls)
- Ensure file is not corrupted

**Permission errors**
- Run as administrator if needed
- Check folder permissions
- Ensure files are not locked by other applications

**Memory issues with large files**
- Close other applications
- Process files in smaller batches
- Increase system virtual memory

**Template generation issues**
- Run `python templates/create_templates.py` manually
- Check pandas and openpyxl installation
- Verify write permissions in project directory

### Performance Tips

- **Large Excel files**: Process in smaller chunks
- **Many folders**: Use batch operations
- **Zip operations**: Adjust compression level based on needs
- **File operations**: Close unnecessary applications
- **UI responsiveness**: Use preview features before large operations

## 🏗️ Development

### Project Structure
```
EruStudio/
├── main.py                    # Main application with modern UI
├── requirements.txt           # Python dependencies
├── setup.py                  # Automated setup and template generation
├── run_erustudio.bat         # Windows batch launcher
├── run_erustudio.ps1         # PowerShell launcher
├── modules/                  # Application modules
│   ├── __init__.py
│   ├── worksheet_sync.py     # Worksheet synchronization
│   ├── bulk_rename.py        # Bulk file renaming
│   ├── multi_zip.py          # Multiple zip creation
│   └── folder_creator.py     # Folder creation
├── templates/                # Excel templates and generator
│   ├── __init__.py
│   ├── create_templates.py   # Template generation script
│   ├── bulk_rename_template.xlsx
│   ├── folder_creator_template.xlsx
│   ├── worksheet_sync_template.xlsx
│   └── multi_zip_template.xlsx
└── README.md                 # This file
```

### Adding New Modules
1. Create a new Python file in the `modules/` directory
2. Implement the module class with required methods
3. Add module button to the main application
4. Update imports and module registration
5. Create corresponding Excel template if needed

### UI Customization
- **Color schemes**: Modify color variables in main.py
- **Layout**: Adjust grid configurations and spacing
- **Typography**: Change font families and sizes
- **Animations**: Enhance hover effects and transitions


## 📥 Download

You can download the latest **EruStudio Windows Installer (.exe)** from the [Releases section](https://github.com/YourUsername/EruStudio/releases).

1. Go to the [Latest Release](https://github.com/YourUsername/EruStudio/releases/latest)
2. Download the file: **EruStudio-Setup.exe**
3. Run the installer and follow the on-screen instructions
4. Launch **EruStudio** from your Start Menu or Desktop shortcut

## 📄 License

This project is provided as-is for educational and personal use.

## 🆘 Support

For issues, questions, or feature requests:
1. Check the troubleshooting section
2. Review error messages carefully
3. Ensure all prerequisites are met
4. Test with sample files first
5. Verify template generation completed successfully

## 📈 Version History

- **v1.1.0**: Modernized UI with dark theme, built-in Excel templates
  - Professional dark color scheme
  - Interactive module cards with hover effects
  - Built-in Excel templates for all modules
  - Responsive design and modern typography
  - Enhanced user experience and visual appeal

- **v1.0.0**: Initial release with four core modules
  - Worksheet synchronization
  - Bulk file renaming
  - Multiple zip creation
  - Folder creation from templates

---


**EruStudio** - Professional File Management Suite for Windows with Modern UI 

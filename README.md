# Active Directory Batch User Creation Tool

[![GitHub License](https://img.shields.io/github/license/menkar85/ADuser_batch_add)](LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/issues)
[![GitHub stars](https://img.shields.io/github/stars/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/stargazers)
[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![Version](https://img.shields.io/badge/version-0.1.3-green.svg)](https://github.com/menkar85/ADuser_batch_add/releases)

A powerful Python-based tool for automating the creation of multiple user accounts in Active Directory. This application reads user data from Excel files and efficiently creates users in your Active Directory domain with comprehensive error handling and logging.

## 🚀 Features

- **📊 Excel Integration**: Import user details from XLSX template files
- **🔄 Batch Processing**: Create multiple user accounts simultaneously
- **🏗️ Automatic OU Creation**: Automatically creates Organizational Units if they don't exist
- **🔧 Customizable Attributes**: Configure user attributes like display name, email, phone, etc.
- **📝 Comprehensive Logging**: Detailed logs for troubleshooting and audit trails
- **🛡️ Error Handling**: Robust error handling with detailed error reporting
- **🌐 Multi-language Support**: Built-in internationalization support
- **🔐 Security**: Supports both LDAP and LDAPS protocols


## 📋 Prerequisites

Before using this tool, ensure you have:

- **Active Directory Environment**: Access to an Active Directory domain
- **Administrative Privileges**: Permissions to create user accounts and OUs
- **Python 3.7+**: Python installed on your system
- **Network Access**: Connectivity to your Active Directory server
- **Excel Template**: Properly formatted Excel file with user data

## 🛠️ Installation

### 1. Clone the Repository

```bash
git clone https://github.com/menkar85/ADuser_batch_add.git
cd ADuser_batch_add
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Prepare Your Excel Template

Copy the provided template file and customize it according to your needs:

```bash
cp "User form template.xlsx" "my_users.xlsx"
```

## 📖 Usage

### 1. Prepare Your Data

Create an Excel file with the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| A | Username (auto-generated) | - |
| B | Surname | Smith |
| C | Password | SecurePass123! |
| D | Full Name | John Smith |
| E | Phone Number | +1-555-0123 |
| F | English Surname (auto-generated) | Smith |
| G | Group/Year | 2024 |
| H | Email (auto-generated) | john.smith@company.com |

### 2. Run the Application

```bash
python application.py
```

### 3. Configure Connection Settings

Fill in the required information in the GUI:

- **LDAP Server**: Your Active Directory server address
- **Username**: Domain administrator username
- **Password**: Domain administrator password
- **Source File**: Path to your Excel file
- **Destination OU**: Target Organizational Unit (e.g., `Users/Staff`)
- **Domain**: Your domain name (e.g., `company.local`)
- **UPN Suffix**: User Principal Name suffix (e.g., `company.local`)
- **Result File**: Output Excel file path
- **Log File**: Log file path
- **Protocol**: LDAP or LDAPS

### 4. Start Import

Click "Start Import" to begin the batch user creation process.

## 📁 File Structure

```
ADuser_batch_add/
├── application.py          # Main application file
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── LICENSE                # License information
├── module/
│   └── forms.py          # GUI form definitions
├── locale/               # Internationalization files
│   ├── en_US/
│   └── ru_RU/
├── testdata/             # Test data directory
└── User form template.xlsx # Excel template
```

## 🔧 Configuration

### Excel Template Format

The application expects an Excel file with the following structure:

- **Row 1**: Headers
- **Row 2+**: User data
- **Columns A-H**: User attributes as described above

### OU Structure

The destination OU should be specified in folder format where "/" represents the hierarchy:

- `Users` - Creates OU in root domain
- `Users/Staff` - Creates Staff OU inside Users OU
- `Departments/IT/Developers` - Creates nested OUs

The application will automatically create any missing OUs in the path.

## 📊 Output Files

### Result Excel File

The application generates a result Excel file with additional columns:

- **Column I**: Success indicator (Y/N)
- **Column J**: Error message (if applicable)

### Log File

Detailed logs are saved to the specified log file, including:

- Connection status
- OU creation/verification
- User creation attempts
- Error details
- Import statistics

## ⚠️ Important Notes

### Duplicate Username Handling

- **First Run**: The application will raise an error for duplicate usernames
- **Subsequent Runs**: Duplicates are marked as existing without errors
- **Recommendation**: For clean imports, delete the destination OU and reimport

### Idempotent Behavior

The application is designed to be idempotent, meaning:
- Existing OUs are reused
- Existing users are skipped
- No duplicate errors on subsequent runs

## 🐛 Troubleshooting

### Common Issues

1. **Connection Failed**
   - Verify LDAP server address
   - Check network connectivity
   - Ensure correct credentials

2. **Permission Denied**
   - Verify administrative privileges
   - Check domain membership
   - Ensure proper OU permissions

3. **Excel File Errors**
   - Verify file format (XLSX)
   - Check column structure
   - Ensure data in required columns

4. **Duplicate Username Errors**
   - Check for existing users
   - Modify usernames in Excel file
   - Delete destination OU for clean import


## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔄 Version History

- **v0.1.3**: Current version with enhanced error handling and logging
- **v0.1.2**: Added multi-language support
- **v0.1.1**: Improved Excel template handling
- **v0.1.0**: Initial release

## ⚡ Performance Tips

- Use LDAPS for secure connections
- Ensure stable network connectivity
- Close other AD management tools during import
- Monitor server resources during large imports

---

**Note**: This tool is designed for administrative use in Active Directory environments. Always test in a non-production environment first and ensure you have proper backups before running batch operations.
# Active Directory Batch User Creation Tool

[![GitHub License](https://img.shields.io/github/license/menkar85/ADuser_batch_add)](LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/issues)
[![GitHub stars](https://img.shields.io/github/stars/menkar85/ADuser_batch_add)](https://github.com/menkar85/ADuser_batch_add/stargazers)
[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![Version](https://img.shields.io/badge/version-0.2.2-green.svg)](https://github.com/menkar85/ADuser_batch_add/releases)

A powerful Python-based tool for automating the creation of multiple user accounts in Active Directory. This application reads user data from Excel files and efficiently creates users in your Active Directory domain with comprehensive error handling, logging, and a modern GUI interface.

## 🚀 Features

- **📊 Excel Integration**: Import user details from XLSX template files
- **🔄 Batch Processing**: Create multiple user accounts simultaneously
- **🏗️ Automatic OU Creation**: Automatically creates Organizational Units if they don't exist
- **🔧 Customizable Attributes**: Configure user attributes like display name, email, phone, etc.
- **📝 Comprehensive Logging**: Detailed logs for troubleshooting and audit trails
- **🛡️ Error Handling**: Robust error handling with detailed error reporting
- **🌐 Multi-language Support**: Built-in internationalization support (English/Russian)
- **🔐 Security**: Supports both LDAP and LDAPS protocols
- **🎨 Modern GUI**: Clean and intuitive ttkbootstrap-based interface
- **🔄 Idempotent Operations**: Safe to run multiple times without creating duplicates

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

| Column | Description | Required | Auto-Generated | Example |
|--------|-------------|----------|----------------|---------|
| A | Username | Yes | Yes | john.smith |
| B | Surname | Yes | No | Smith |
| C | Password | Yes | No | SecurePass123! |
| D | Full Name | Yes | No | John Smith |
| E | Phone Number | No | No | +1-555-0123 |
| F | English Surname | Yes | Yes | Smith |
| G | Group/Year | No | No | 2024 |
| H | Email | Yes | Yes | john.smith@company.com |

### 2. Run the Application

```bash
python application.py
```

### 3. Configure Connection Settings

Fill in the required information in the GUI:

- **LDAP Server**: Your Active Directory server address (e.g., `dc.company.local`)
- **Username**: Domain administrator username (e.g., `administrator@company.local`)
- **Password**: Domain administrator password
- **Source File**: Path to your Excel file
- **Destination OU**: Target Organizational Unit (e.g., `Users/Staff`)
- **Domain**: Your domain name (e.g., `company.local`)
- **UPN Suffix**: User Principal Name suffix (e.g., `company.local`)
- **Result File**: Output Excel file path and name for results
- **Log File**: Log file path and name for detailed logs
- **Protocol**: LDAP or LDAPS (recommended for security)

### 4. Start Import

Click "Start Import" to begin the batch user creation process. The application will:

1. Connect to Active Directory
2. Create destination OUs if they don't exist
3. Process each user in the Excel file
4. Generate a result file with success/failure status
5. Create detailed logs

## 📁 Project Structure

```
ADuser_batch_add/
├── application.py              # Main application file
├── requirements.txt            # Python dependencies
├── README.md                  # This file
├── LICENSE                    # License information
├── data.pkl                   # Application state persistence
├── module/
│   └── forms.py              # GUI form definitions
├── locale/                    # Internationalization files
│   ├── en_US/
│   │   └── LC_MESSAGES/
│   └── ru_RU/
│       └── LC_MESSAGES/
└── User form template.xlsx    # Excel template
```

## 🔧 Configuration

### Excel Template Format

The application expects an Excel file with the following structure:

- **Row 1**: Headers (Username, Surname, Password, etc.)
- **Row 2+**: User data
- **Columns A-H**: User attributes as described above

### OU Structure

The destination OU should be specified in folder format where "/" represents the hierarchy:

- `Users` - Creates OU in root domain
- `Users/Staff` - Creates Staff OU inside Users OU
- `Departments/IT/Developers` - Creates nested OUs

The application will automatically create any missing OUs in the path.

### Auto-Generated Fields

The following fields are automatically generated if not provided:

- **Username**: Generated from surname and given name
- **Email**: Generated from username and domain
- **English Surname**: Transliterated from original surname

## 📊 Output Files

### Result Excel File

The application generates a result Excel file with additional columns:

- **Column I**: Success indicator (Y/N)
- **Column J**: Error message (if applicable)

### Log File

Detailed logs are saved to the specified log file, including:

- Connection status and authentication
- OU creation/verification
- User creation attempts and results
- Error details and stack traces
- Import statistics and summary

## ⚠️ Important Notes

### Duplicate Username Handling

- **First Run**: The application will raise an error for duplicate usernames
- **Subsequent Runs**: Duplicates are marked as existing without errors
- **Recommendation**: For clean imports, delete the destination OU and reimport

### Idempotent Behavior

The application is designed to be idempotent, meaning:

- Existing OUs are reused (not recreated)
- Existing users are skipped (not recreated)
- No duplicate errors on subsequent runs
- Safe to run multiple times

### Security Considerations

- Use LDAPS protocol for secure connections
- Store credentials securely
- Test in non-production environment first
- Ensure proper backups before batch operations

## 🐛 Troubleshooting

### Common Issues

1. **Connection Failed**
   - Verify LDAP server address and port
   - Check network connectivity to AD server
   - Ensure correct domain credentials
   - Verify firewall settings
   - If using LDAPS ensure proper configuration of AD

2. **Permission Denied**
   - Verify administrative privileges on domain
   - Check domain membership of executing machine
   - Ensure proper OU permissions
   - Verify account lockout status

3. **Excel File Errors**
   - Verify file format (XLSX only)
   - Check column structure matches template
   - Ensure required data in mandatory columns
   - Validate data types (no special characters in usernames)

4. **Duplicate Username Errors**
   - Check for existing users in AD
   - Modify usernames in Excel file
   - Delete destination OU for clean import

### Performance Tips

- Use LDAPS for secure connections
- Ensure stable network connectivity
- Close other AD management tools during import
- Monitor server resources during large imports
- Process users in smaller batches for large datasets

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📞 Support

If you encounter any issues or have questions:

1. Check the troubleshooting section above
2. Review the log files for detailed error information
3. Open an issue on GitHub with detailed information

---

**⚠️ Disclaimer**: This tool is designed for administrative use in Active Directory environments. Always test in a non-production environment first and ensure you have proper backups before running batch operations. The authors are not responsible for any data loss or system issues that may occur during use.
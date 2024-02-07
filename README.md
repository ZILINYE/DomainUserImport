# AD User Import Script

This PowerShell script is designed for use in educational settings, specifically colleges, to automate the creation of student accounts in Active Directory (AD). It utilizes data from an Excel file or a MySQL database to generate user accounts efficiently.

## Functionality

- **Import AD Users**: Automatically creates AD user accounts based on a predefined Excel file format or data fetched from a MySQL database.
- **Group and OU Assignment**: Dynamically assigns users to specific groups and Organizational Units (OU) based on the data provided.
- **Password Assignment**: Sets user passwords by combining their birthday and student ID for initial login security.
- **Password Output**: Generates a text file containing the user passwords for administrative records.
- **Change Log**: Outputs a log file detailing all changes made within the Domain Controller (DC) during the script execution.

## Usage

1. Ensure you have the necessary permissions to create and manage AD user accounts within your network.
2. Prepare the Excel file or MySQL database with student information according to the required format.
3. Modify the script to connect to your specific Excel source or MySQL database, specifying the path or connection details.
4. Customize the group, OU assignments, and password generation logic as per your institutional policies.
5. Run the script in PowerShell with appropriate administrative privileges:

```powershell
.\ADUserImport.ps1
```
## Requirements

- PowerShell 5.1 or higher
- Active Directory PowerShell module
- Access to AD with sufficient privileges to create and manage user accounts
- Excel file or MySQL database containing student information

## Note
This script should be used responsibly and with consideration of organizational policies and data protection laws. Ensure that passwords are handled securely and that students are instructed to change their passwords upon first login.

## Disclaimer

This script is provided "as is", without warranty of any kind. Use at your own risk. Always test scripts in a non-production environment before deploying them into production.

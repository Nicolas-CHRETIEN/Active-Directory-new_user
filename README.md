README - New User Creation Tool
===============================

Overview
--------
This tool is a graphical PowerShell application that allows an administrator to create a new Active Directory user and configure a personal folder with appropriate permissions.

Features
--------
- Graphical input for first name, last name, password, and folder location
- Automatic generation of account name and email
- OU selection and optional creation of new Organizational Units
- Group membership assignment with optional creation of new groups
- Creation and sharing of a personal folder (SMB Share)
- Assignment of exclusive NTFS permissions to the user
- Fully silent execution when compiled as .EXE

How to Use
----------
1. Double-click `new_user.exe` to launch the tool.
2. Enter the user’s first and last name when prompted.
3. Choose whether to assign the user to an Organizational Unit (OU). You may create one if needed.
4. Set a secure password for the new account.
5. Choose whether to add the user to an existing group. You may create one if needed.
6. Select the location for the user's personal folder.
7. The script will automatically:
   - Create the user in AD
   - Create and share the folder
   - Assign proper NTFS and SMB permissions

Technical Notes
---------------
- This tool requires PowerShell and Active Directory modules.
- You must run it with sufficient privileges (Domain Admin or delegated rights).
- The compiled EXE includes a custom icon (`new_user.ico`).
- The folder permissions use `icacls` to grant full access only to the user.

Author
------
Nicolas Chrétien



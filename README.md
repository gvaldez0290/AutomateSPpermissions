# AutomateSPpermissions
Shows how to a complex SharePont permissions schema based on groups outlined in a csv

This README provides an overview and usage instructions for a PowerShell script designed to copy files and folders from one SharePoint library to another while maintaining the folder structure and permissions. This script is useful for scenarios where you need to replicate a SharePoint library's content to another location.

## Prerequisites
Before using this script, make sure you have the following prerequisites in place:

PowerShell 7: Ensure you have PowerShell 7 installed on your machine.

SharePoint Online: You should have access to the source and destination SharePoint Online libraries. This script is designed for SharePoint Online environments.

PnP PowerShell: Install the PnP PowerShell module using the following command:

```powershell
Import-Module PnP.PowerShell -Force
```
## Usage
Follow these steps to use the script:

Open the script file in a text editor or PowerShell Integrated Scripting Environment (ISE).

Update the following variables in the script with your specific values:

-Username: Your SharePoint Online username.
-Password: Your SharePoint Online password.
-SiteUrl: The URL of your SharePoint Online site.
-SourceLibrary: The name of the source library from which you want to copy files and folders.
-DestinationLibrary: The name of the destination library where you want to copy the files and folders.

Run the script in PowerShell. It will prompt you to enter the name of the destination library using a graphical user interface (GUI). Enter the name and click "OK."

The script will copy files and folders from the source library to the destination library while preserving the folder structure.

After copying, the script sets permissions based on a CSV file containing folder IDs, group names, and permission levels. Make sure to update the CSV file path and structure as needed.

The script will also add a list item to the destination library, which you can customize to your requirements.

Finally, it will launch your web browser to confirm that the Document library was created correctly.

## Desktop Shortcut
Create a destop shortcut by creating the following .bat file. 
```bat
"C:\Program Files\PowerShell\7\pwsh.exe" -File "path.ps1"
```
Create a shortcut of this file and paste it onto your desktop. This allows you to run the program with the click of a button.

## Notes
The script uses the PnP PowerShell module for SharePoint operations, so ensure it is installed and configured properly.
Make sure to handle sensitive information like passwords securely, possibly using SharePoint app passwords or other secure methods.
Disclaimer: Use this script with caution, especially in production environments, and test it thoroughly in a safe environment before deploying it to ensure it meets your specific requirements and does not inadvertently affect your SharePoint data.


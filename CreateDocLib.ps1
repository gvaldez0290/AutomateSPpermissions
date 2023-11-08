Import-Module PnP.PowerShell -Force

#Cred-Variables
$Username = "username@email.com"
$Password = "password"

# Variables
$SiteUrl = "https://[orgsitename].sharepoint.com/sites/[sitename]/"
$SourceLibrary = "[template library name]"

$ConfirmPreference = 'None'

# GUI 
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterScreen"
$form.Text = "Copy SharePoint Library"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(30,20)
$label.Size = New-Object System.Drawing.Size(400,20)
$label.Text = "Please enter the name of the destination library:"
$form.Controls.Add($label)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(30,50)
$textBox.Size = New-Object System.Drawing.Size(340,20)
$form.Controls.Add($textBox)
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(30,90)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = "OK"
$okButton.Add_Click({
    $form.Tag = $textBox.Text
    $form.Close()
})
$form.Controls.Add($okButton)
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(120,90)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = "Cancel"
$cancelButton.Add_Click({
    $form.Tag = $null
    $form.Close()
})
$form.Controls.Add($cancelButton)
$form.ShowDialog()

# Use the destination library name entered by the user
$DestinationLibrary = $form.Tag

# Hash table to keep track of processed items
$processedItems = @{}

# Connect to SharePoint Online site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Create New Document Library with the same structure
New-PnPList -Title $DestinationLibrary -Template DocumentLibrary


# Function to recursively copy files, and folders
Function Copy-FilesAndFolders {
    param(
        [string]$SourceFolder,
        [string]$DestFolder
    )

    # Get items in the source folder
    $items = Get-PnPListItem -List $SourceLibrary -FolderServerRelativeUrl $SourceFolder

    # Loop through each item
    foreach ($item in $items) {
        try {
            $isFolder = $item.FileSystemObjectType -eq "Folder"
            $itemUrl = $item.FieldValues["FileRef"]
            $itemName = $item.FieldValues["FileLeafRef"]

            # Skip this item if it has already been processed
            if ($processedItems.ContainsKey($itemUrl)) {
                continue
            }

            # Mark this item as processed
            $processedItems[$itemUrl] = $true

            # Create the folder/file in the destination within "9999" folder
            $newDestFolder = "/sites/[sitename]/$DestinationLibrary/9999/$DestFolder/$itemName"

            if ($isFolder) {
                Add-PnPFolder -Name $itemName -Folder $newDestFolder
            } else {
                Copy-PnPFile -SourceUrl $itemUrl -TargetUrl $newDestFolder -OverwriteIfAlreadyExists -Force
            }

            # If it's a folder, recursively copy its contents
            if ($isFolder) {
                Copy-FilesAndFolders -SourceFolder $itemUrl -DestFolder $newDestFolder
            }
        } catch {
            # Ignore "File Not Found" errors and continue
            if ($_.Exception.Message -notmatch "File Not Found") {
                Write-Host $_.Exception.Message
            }
        }
    }
}

# Start the recursive copy, all subsequent folders will be copied into "9999"
Copy-FilesAndFolders -SourceFolder "/sites/[sitename]/$SourceLibrary" -DestFolder ""


#####################################
## This section sets the permissions based on subfolder ID numbers by reading the CSV file
# Connect to SharePoint site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Clear permissions for the list item with ID 1
Set-PnPListItemPermission -User $Username -List $DestinationLibrary -Identity 1 -ClearExisting

# Output message
Write-Host "Permissions cleared for folder with ID 1."

#####read csv and set permissions######
$csvData = Import-Csv -Path "path.csv"

# Loop through each row
foreach ($row in $csvData) {
    # Extract the necessary details
    $folderId = $row.FolderID
    $groupName = $row.GroupName
    $permission = $row.Permission

    # Check for the type of permission and set it accordingly
    switch ($permission) {

        "Contribute" {
            Set-PnPListItemPermission -List $DestinationLibrary -Identity $folderId -Group $groupName -AddRole "Contribute"
        }
        "Full Control" {
            Set-PnPListItemPermission -List $DestinationLibrary -Identity $folderId -Group $groupName -AddRole "Full Control"
        }
        "Edit" {
            Set-PnPListItemPermission -List $DestinationLibrary -Identity $folderId -Group $groupName -AddRole "Edit"
        }
        "Read" {
            Set-PnPListItemPermission -List $DestinationLibrary -Identity $folderId -Group $groupName -AddRole "Read"
        }
    }

    Write-Host "Permissions set for folder with ID $folderId for group $groupName with $permission permission."
}


# Debugging output
Write-Host "List Name: $ListName"
Write-Host "New Title: $NewTitle"
Write-Host "Link URL: $NewProjectFolderLinkUrl"


# Add list item with specified content type
Add-PnPListItem -List $ListName -Values @{
    "Title" = $NewTitle;
    "ProjectFolder" = $NewProjectFolderLinkUrl   
}

Write-Host "New list item created in the SharePoint list!"


# Disconnect
Disconnect-PnPOnline

Write-Host "Folders and files copied successfully!"

#launch the browser to ensure the Document library was created correctly
Start-Process "chrome.exe" "https://[orgsitename].sharepoint.com/sites/[sitename]/_layouts/15/viewlsts.aspx?view=14"


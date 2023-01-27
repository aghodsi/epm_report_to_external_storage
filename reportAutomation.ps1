Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Installing the Graph Module
If (-not(Get-InstalledModule "Microsoft.Graph" -ErrorAction silentlycontinue)) {
   Write-Host "Graph API is not installed. Installing."
   $ProgressPreference = "SilentlyContinue"
   Install-Module "Microsoft.Graph" -Scope CurrentUser
}
Else {
   # Updates disabled, enable if you like 
   Write-Host "Module exists. Updating it"
   # Update-Module "Microsoft.Graph"
}

Import-Module Microsoft.Graph.Mail
# Read emails to get URLs and extract download link. User read all permissions to enhance the data with additional data from AAD/ Files.ReadWrite to upload files to sharepoint
Connect-MgGraph -Scopes "Mail.Read", "User.Read.All", 'Sites.ReadWrite.All', 'Files.ReadWrite.All'

#replace with your own username
$upn = '<<replace me with UPN>>'

# Filter messages from EPM and those that are reports. Filter then where they are from the last 7 days. Change it to number of days you'd like to go back
$dateFilter = $(Get-Date).AddDays(-1) 
$reportMessages = Get-MgUserMessage -UserId $upn -Filter "from/emailAddress/address eq 'no-reply-ciem@microsoft.com' and contains(subject,'report')"  -All | Where-Object { $_.ReceivedDateTime -gt $dateFilter }

$pattern = "\b(https?:\/\/[^\s]+download)\b.+?\b(All Files Zip)\b"
foreach ($message in $reportMessages) {
   if ($message.Body.Content -match $pattern) {
      # Print the matched URL
      Write-Host "Found URL: $($matches[1]). Opening URL"
      # Opens URL and Downloads content, Sleeping 5 seconds after each URL
      Start-Process $($matches[1]) -WindowStyle Minimized
      Start-Sleep -Seconds 5
   }

}

function Get-FileName {
   <#
.SYNOPSIS
   Show an Open File Dialog and return the file selected by the user

.DESCRIPTION
   Show an Open File Dialog and return the file selected by the user

.PARAMETER WindowTitle
   Message Box title
   Mandatory - [String]

.PARAMETER InitialDirectory
   Initial Directory for browsing
   Mandatory - [string]

.PARAMETER Filter
   Filter to apply
   Optional - [string]

.PARAMETER AllowMultiSelect
   Allow multi file selection
   Optional - switch

 .EXAMPLE
   Get-FileName
    cmdlet Get-FileName at position 1 of the command pipeline
    Provide values for the following parameters:
    WindowTitle: My Dialog Box
    InitialDirectory: c:\temp
    C:\Temp\42258.txt

    No passthru paramater then function requires the mandatory parameters (WindowsTitle and InitialDirectory)

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp
   C:\Temp\41553.txt

   Choose only one file. All files extensions are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect
   C:\Temp\8544.txt
   C:\Temp\42258.txt

   Choose multiple files. All files are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "text file (*.txt) | *.txt"
   C:\Temp\AES_PASSWORD_FILE.txt

   Choose multiple files but only one specific extension (here : .txt) is allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "Text files (*.txt)|*.txt| csv files (*.csv)|*.csv | log files (*.log) | *.log"
   C:\Temp\logrobo.log
   C:\Temp\mylogfile.log

   Choose multiple file with the same extension

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "selected extensions (*.txt, *.log) | *.txt;*.log"
   C:\Temp\IPAddresses.txt
   C:\Temp\log.log

   Choose multiple file with different extensions
   Nota :It's important to have no white space in the extension name if you want to show them

.EXAMPLE
 Get-Help Get-FileName -Full

.INPUTS
   System.String
   System.Management.Automation.SwitchParameter

.OUTPUTS
   System.String

.NOTESs
  Version         : 1.0
  Author          : O. FERRIERE
  Creation Date   : 11/09/2019
  Purpose/Change  : Initial development

  Based on different pages :
   mainly based on https://blog.danskingdom.com/powershell-multi-line-input-box-dialog-open-file-dialog-folder-browser-dialog-input-box-and-message-box/
   https://code.adonline.id.au/folder-file-browser-dialogues-powershell/
   https://thomasrayner.ca/open-file-dialog-box-in-powershell/
#>
   [CmdletBinding()]
   [OutputType([string])]
   Param
   (
      # WindowsTitle help description
      [Parameter(
         Mandatory = $true,
         ValueFromPipelineByPropertyName = $true,
         HelpMessage = "Message Box Title",
         Position = 0)]
      [String]$WindowTitle,

      # InitialDirectory help description
      [Parameter(
         Mandatory = $true,
         ValueFromPipelineByPropertyName = $true,
         HelpMessage = "Initial Directory for browsing",
         Position = 1)]
      [String]$InitialDirectory,

      # Filter help description
      [Parameter(
         Mandatory = $false,
         ValueFromPipelineByPropertyName = $true,
         HelpMessage = "Filter to apply",
         Position = 2)]
      [String]$Filter = "All files (*.*)|*.*",

      # AllowMultiSelect help description
      [Parameter(
         Mandatory = $false,
         ValueFromPipelineByPropertyName = $true,
         HelpMessage = "Allow multi files selection",
         Position = 3)]
      [Switch]$AllowMultiSelect
   )

   # Load Assembly
   Add-Type -AssemblyName System.Windows.Forms

   # Open Class
   $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

   # Define Title
   $OpenFileDialog.Title = $WindowTitle

   # Define Initial Directory
   if (-Not [String]::IsNullOrWhiteSpace($InitialDirectory)) {
      $OpenFileDialog.InitialDirectory = $InitialDirectory
   }

   # Define Filter
   $OpenFileDialog.Filter = $Filter

   # Check If Multi-select if used
   if ($AllowMultiSelect) {
      $OpenFileDialog.MultiSelect = $true
   }
   $OpenFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
   $OpenFileDialog.ShowDialog() | Out-Null
   if ($AllowMultiSelect) {
      return $OpenFileDialog.Filenames
   }
   else {
      return $OpenFileDialog.Filename
   }
}

$dlFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$files = Get-FileName -InitialDirectory $dlFolder -WindowTitle 'Select downloaded zip files' -Filter 'Compressed Files (*.zip)|*.zip' -AllowMultiSelect

# SPO site where we are going to save the files to. Chang the below Search Query to match your SharePoint Site that you created
$siteId = $(Get-MgSite -Search "<<replace me with site name>>" | Select-Object Id).Id
#Add Days is for testing reasons. You can change it to test your folder creation
$today = $(Get-Date).AddDays(0).ToString('yyyy-MM-dd')

$now = $(Get-Date -Format 'HH_mm')
$requestBody = @{
   "name" = $today+"_"+$now
   "folder" = @{}
   } | ConvertTo-Json

$folder = $(Invoke-MgGraphRequest -Method "POST" -Uri "v1.0/sites/${siteId}/drive/root/children" -Body $requestBody -ContentType "application/json")


foreach ($file in $files) {
   #unzip
   #run through CSV and enhance 
   #upload to SPO
   Write-Host "File: ${file}"
   $CloudEnvNameStartDelimiterIndex = $file.LastIndexOf("-")
   $CloudEnvNameEndDelimiterIndex = $file.LastIndexOf(".")
   $CloudEnvironment = $file.Substring($CloudEnvNameStartDelimiterIndex+1, ($CloudEnvNameEndDelimiterIndex-$CloudEnvNameStartDelimiterIndex-1))
   $fileList = $(Expand-Archive -Path $file -PassThru -Force)
   if ($file.Contains('User_Entitlements_And_Usage')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/user_eau_summary_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*summary*'} | Select-Object FullName).FullName -ContentType "text/csv"
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/user_eau_details_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*details*'} | Select-Object FullName).FullName -ContentType "text/csv"
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/user_eau_permissions_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*permissions*'} | Select-Object FullName).FullName -ContentType "text/csv"
   }
   if ($file.Contains('Permissions_Analytics_Report')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/PAR_${CloudEnvironment}_${today}.xlsx:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.xlsx' -and $_.DirectoryName -Like '*reports*'} | Select-Object FullName).FullName -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   }
   if($file.Contains('Identity_Permissions')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/identity_perms_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*reports*'} | Select-Object FullName).FullName -ContentType "text/csv"
   }
   if($file.Contains('PCI_History')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/PCI_History_${CloudEnvironment}_${today}.xlsx:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.xlsx' -and $_.DirectoryName -Like '*reports*'} | Select-Object FullName).FullName -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   }
   if($file.Contains('Group_Entitlements_And_Usage')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/group_summary_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*summary*'} | Select-Object FullName).FullName -ContentType "text/csv"
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/group_memberships_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*memberships*'} | Select-Object FullName).FullName -ContentType "text/csv"
   }
   if($file.Contains('All_Permissions_for_Identity')){
      Invoke-MgGraphRequest -Method "PUT" -Uri "v1.0/sites/${siteId}/drive/items/root:/$($folder.name)/perms_for_specific_identities_${CloudEnvironment}_${today}.csv:/content" -InputFilePath $($fileList | where-Object {$_.Extension -eq '.csv' -and $_.DirectoryName -Like '*reports*'} | Select-Object FullName).FullName -ContentType "text/csv"
   }
}

Disconnect-MgGraph
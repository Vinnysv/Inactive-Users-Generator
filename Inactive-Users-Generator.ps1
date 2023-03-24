# Script to sort users based on last login date
# Vincent Spagnola
# First created 10-17-2022
# Last edited 10-17-2022

# Import Dependencies, AzureAD to grab data and ImportExcel to sort and format data.
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath

# Getting directory path of script regardless of PowerShell directory
$year =  Get-Date -Format 20%y
$month = Get-Date -Format %M
$day = Get-Date -Format %d
$logs_path = $args -join " "
$previous_month = $month -1
$date = $year+ "-" + $month + "-" + $day
$previous_date = $year+ "-" + $previous_month + "-" + $day
$finaldir = $dir + "\Users(" + $date + ").xlsx"

if ([string]::IsNullOrEmpty($logs_path)){
    $logs_path = Read-Host -Prompt 'Input Path To Spreadsheet' # Input for user to put path
}

# If not admin, restart as admin
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList ($arguments + $logs_path)
    Break
}

# Install ImportExcel module
Install-Module ImportExcel

# Convert CSV to XLSX if needed
if ($logs_path.contains(".csv")){
    $logs_path_new = $logs_path.replace(".csv",".xlsx")
    Import-CSV ($logs_path) | Export-Excel ($logs_path_new)
    $logs_path = $logs_path_new
}

# Import last exported data to arrays and sort them by column
$logs = Import-Excel -Path $logs_path
$logs_names = $logs."Display Name"
$logs_teams = $logs."Teams Last Activity Date"
$logs_sharepoint = $logs."SharePoint Last Activity Date"
$logs_onedrive = $logs."OneDrive Last Activity Date"
$logs_exchange = $logs."Exchange Last Activity Date"

# Initialize an empty list to store data
$data = @()

# Iterate through each name in the imported Azure data
$i = 0
foreach ($user in $logs_names) {
    # Replace empty values with a default date
    if ([string]::IsNullOrEmpty($logs_teams[$i])) {
        $logs_teams[$i] = 1991
    }
    if ([string]::IsNullOrEmpty($logs_sharepoint[$i])) {
        $logs_sharepoint[$i] = 1991
    }
    if ([string]::IsNullOrEmpty($logs_exchange[$i])) {
        $logs_exchange[$i] = 1991
    }
    if ([string]::IsNullOrEmpty($logs_onedrive[$i])) {
        $logs_onedrive[$i] = 1991
    }

    # Check if the last activity dates are older than the previous_date
    if ((get-date $logs_teams[$i]) -lt (get-date $previous_date) -and ((get-date $logs_sharepoint[$i]) - lt (get-date $previous_date)) -and ((get-date $logs_onedrive[$i]) -lt (get-date $previous_date)) -and ((get-date $logs_exchange[$i]) -lt (get-date $previous_date))) {
		# Create a new object to store user data
		$userData = New-Object PSObject -Property @{
			'User' = $user
			'Last Teams Login Date' = if ($logs_teams[$i] -eq 1991) { "" } else { $logs_teams[$i] }
			'Last OneDrive Login Date' = if ($logs_onedrive[$i] -eq 1991) { "" } else { $logs_onedrive[$i] }
			'Last Exchange Login Date' = if ($logs_exchange[$i] -eq 1991) { "" } else { $logs_exchange[$i] }
			'Last Sharepoint Login Date' = if ($logs_sharepoint[$i] -eq 1991) { "" } else { $logs_sharepoint[$i] }
		}
		# Add the new user data object to the data list
		$data += $userData
}
$i++

#Export final data to an Excel file
$data | Export-Excel $finaldir -AutoSize -BoldTopRow -WorkSheetName "Users"

#Remove the original file
Remove-Item -path $logs_path
<#
.SYNOPSIS
    User Logon Report from Horizon Events DB
.DESCRIPTION
    User Logon Report from Horizon Events DB
    
    Requires SQL Server Module and ImportExcel Module

.NOTES
    Version:          1.0.0
    Author:           Chris Hildebrandt
    GitHub:           https://github.com/childebrandt42
    Twitter:          @childebrandt42
    Date Created:     10/21/2023
    Date Updated:     10/21/2023
 

#>

#---------------------------------------------------------------------------------------------#
#                                  Script Varribles                                           #
#---------------------------------------------------------------------------------------------#

# SQL Account info
$Creds = Get-Credential -Message "Enter Local SQL Account"

# Varible for how many days to look back
$Days = 365

# How would you like the report
$ReportType = "Excel" # CSV or Excel

# Include Pivot Table and Chart NOTE: This only works on Excel reports
$PivotTable = $true # $true or $false

# SQL Server Name
$SQLServers = @('SQL Server FQDN') # SQL Server FQDN

# DB Name
$dbName = "Horizon Envents DB Name" # Horizon Events DB Name

# Report Name
$ReportName = "UserReport" # Report Name

# Report Save Location
$ReportSaveLocation = "C:\Reports\Usage"

# Days to look back
$TimeBack = get-date -date $(get-date).adddays(-$Days) -format "yyyy-MM-dd HH:mm:ss"

#---------------------------------------------------------------------------------------------#
#                                  Powershell Modeuls                                         #
#---------------------------------------------------------------------------------------------#

# If SQL Server Module is not installed, install it
if(-not (Get-Module sqlserver -ListAvailable)){
    Write-Host "SQL Server Module is not installed, installing now"
    Install-Module sqlserver -Scope CurrentUser -Force
}

If($ReportType -eq "Excel"){
    Write-Host "Report Type is Excel"
    # If ImportExcel Module is not installed, install it
    if(-not (Get-Module ImportExcel -ListAvailable)){
        Write-Host "ImportExcel Module is not installed, installing now"
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
}

# Import Modules
Import-Module ImportExcel
Import-Module sqlserver

#---------------------------------------------------------------------------------------------#
#                                  Script Logic                                               #
#---------------------------------------------------------------------------------------------#

# Set Date for Log
$ScriptDateLog = Get-Date -Format MM-dd-yyy-HH-mm-ss

# If Report Save Location does not exist, create it
if (-not (Test-Path $ReportSaveLocation)) {
    Write-Host "Creating Report Save Location"
    New-Item -Path $ReportSaveLocation -ItemType Directory
}

# Function to retrieve events from SQL Server
function Get-Events {
    param (
        [string]$SQLServer,
        [string]$SQLQuery
    )
    Invoke-Sqlcmd -Credential $Creds -ServerInstance $SQLServer -Database $dbName -Query $SQLQuery -TrustServerCertificate | Select-Object ModuleAndEventText, Time, Node, DesktopId
}

# Create Blank Arrays
$Events = @('')
$EventsHST = @('')

$SQLQueryHST = "SELECT * from $dbName.event_historical where (EventType = 'AGENT_CONNECTED') and (Time > '$TimeBack') order by time desc"
$SQLQueryEvent = "SELECT * from $dbName.event where (EventType = 'AGENT_CONNECTED') and (Time > '$TimeBack') order by time desc"

foreach($SQLServer in $SQLServers){
    Write-Host "Processing $SQLServer for Historical Events"
    $EventsHST += Get-Events -SQLServer $SQLServer -SQLQuery $SQLQueryHST
    Write-Host "Processing $SQLServer for Events"
    $Events += Get-Events -SQLServer $SQLServer -SQLQuery $SQLQueryEvent
}

# Create Blank Arrays
$EventData = @('')
$EventDataHST = @('')
$UsernameHSTEDT = @('')
$EventHST = @('')

# Create Event Data Array
Foreach ($Event in $Events){
    $EventData += [pscustomobject]@{
        NodeID = $Event.Node
        PoolID = $Event.DesktopId
    }
}

# Create Historical Event Data Array
Foreach ($EventHST in $EventsHST){
    if($EventHST){
        $UsernameHSTEDT = ''
        $UsernameHSTEDT = $EventHST.ModuleAndEventText | Out-String
        $UsernameHSTEDT = $UsernameHSTEDT.Trim('User ')
        $UsernameHSTEDT = $UsernameHSTEDT.substring(0,$UsernameHSTEDT.IndexOf(' '))
        $UsernameHSTEDT = $UsernameHSTEDT.Split('\')[$($UsernameHSTEDT.Split('\').Count-1)]
        $UsernameHSTEDT = $UsernameHSTEDT.Split('\')[$($UsernameHSTEDT.Split('\').Count-1)]
        $DesktopIDHST = $EventData | Where-Object {$_.NodeID -like $EventHST.Node} | Select-Object -First 1 | Select-Object PoolID

        # Create Historical Event Data Array to Export
        if (-not ([string]::IsNullOrEmpty($EventHST)))
        {
            $DateTime = [DateTime]::Parse($EventHST.Time)
            
            $EventDataHST += [pscustomobject]@{
                UserName = $UsernameHSTEDT
                LogonTime = $DateTime
                NodeID = $EventHST.Node
                PoolID = $DesktopIDHST.PoolID
            }
        }
    }

}

# Remove Blank Lines from Array
$EventDataHST = $EventDataHST | Where-Object {$_.UserName -ne $null}

# Create Report
If ($ReportType -eq "CSV"){
    Write-Host "Exporting to CSV"
    # Export to Excel as CSV
    $EventDataHST | Select-Object ('UserName','LogonTime','NodeID','PoolID') | Export-Csv -Path "$ReportSaveLocation\$ReportName-$ScriptDateLog.csv" -NoTypeInformation
}

# Create Report
If ($ReportType -eq "Excel" -and $PivotTable -eq $false){
    Write-Host "Exporting to Excel without Pivot Table"
    # Export to Excel as XLSX
    $EventDataHST | Select-Object ('UserName','LogonTime','NodeID','PoolID') | Export-Excel -Path "$ReportSaveLocation\$ReportName-$ScriptDateLog.xlsx" -WorksheetName "LogonData" -ClearSheet -AutoSize -BoldTopRow
}

# Create Report
if ($ReportType -eq "Excel" -and $PivotTable -eq $true) {
    Write-Host "Exporting to Excel with Pivot Table"
    # Export to Excel as XLSX
    $EventDataHST | Select-Object ('UserName','LogonTime','NodeID','PoolID') | Export-Excel -Path "$ReportSaveLocation\$ReportName-$ScriptDateLog.xlsx" -WorksheetName "LogonData" -ClearSheet -AutoSize -BoldTopRow -IncludePivotTable -PivotTableName "User Logon Report Table" -PivotRows "LogonTime","PoolID","NodeID" -StartRow 2 -PivotData @{"UserName"="Count"} -IncludePivotChart -PivotChartType CylinderCol
   
}
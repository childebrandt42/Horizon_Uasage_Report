<p align="center">
    <a href="https://twitter.com/childebrandt42" alt="Twitter">
            <img src="https://img.shields.io/twitter/follow/Childebrandt42.svg?style=social"/></a>
</p>
<!-- ********** DO NOT EDIT THESE LINKS ********** -->

# VMware Horizon Usage Report

VMware Horizon Usage report which works in conjunction with [SQLServer Powershell Module](https://www.powershellgallery.com/packages/SqlServer/22.1.1) and [ImportExcel Powershell Module](https://github.com/dfinke/ImportExcel).

Please refer to my blog [website](https://www.childebrandt42.blog) for more detailed information about this project.

# :books: Sample Reports

## Sample Report

Sample Horizon Usage Report CSV format: [Sample-UserReport-CSV.csv](https://htmlpreview.github.io/?https://raw.githubusercontent.com/childebrandt42/Horizon_Uasage_Report/main/Samples/Sample-UserReport-CSV.csv)

Sample Horizon Usage Report Excel format without Pivot Table: [Sample-UserReport-No-PivotTable.xlsx](https://htmlpreview.github.io/?https://raw.githubusercontent.com/childebrandt42/Sample-UserReport-No-PivotTable.xlsx)

Sample Horizon Usage Report CSV format: [Sample-UserReport-PivotTable.xlsx](https://htmlpreview.github.io/?https://raw.githubusercontent.com/childebrandt42/Horizon_Uasage_Report/main/Samples/Sample-UserReport-PivotTable.xlsx)

# :beginner: Getting Started
Below are the instructions on how to run the VMware Horizon Usage Report

### PowerShell
This report is compatible with the following PowerShell versions;

<!-- ********** Update supported PowerShell versions ********** -->
| Windows PowerShell 5.1 |     PowerShell 7    |
|:----------------------:|:--------------------:|
|   :white_check_mark:   | :white_check_mark: |
## :wrench: System Requirements
<!-- ********** Update system requirements ********** -->
PowerShell 5.1 or PowerShell 7, and the following PowerShell modules are required for generating a VMware Horizon Usage Report.

- [VMware PowerCLI Module](https://www.powershellgallery.com/packages/VMware.PowerCLI/)
- [SQL Server Module](https://www.powershellgallery.com/packages/SqlServer/)
- [Import Excel Module](https://www.powershellgallery.com/packages/ImportExcel/)

## :package: Instructions

Download script

Place script on machine that can talk to Horizon Events DB SQL server

Fill in the Varribles:  
* $Creds = Credentials for SQL Server that has read rights to the Horizon Events Database. Must be local SQL Credentials
* $Days =  Number of Days you want the report to go back. 
* $ReportType = Would you like the report to be in Excel or CSV format? 
* $PivotTable = Enabe Pivot Table or not
* $SQLServers = SQL Server FQDN
* $dbName = Horizon Events DB Name
* $ReportName = Report Name
* $ReportSaveLocation = Report Save Location

Then Run the script. 

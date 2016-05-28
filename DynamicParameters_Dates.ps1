#Requires -Version 3
Set-StrictMode -Version Latest

Add-Type -assembly "System.IO.Compression"
Add-Type -assembly "System.IO.Compression.Filesystem"

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
# Updated by Jake Minette @ GoDaddy for workbook    #
# parameter updates.                                #
# email: jminette@godaddy.com                       #
#####################################################

#Initial variables
$myTableauServer     = "localhost:8000"
$myTableauServerUser = "TableauAdmin"
$myTableauServerPW   = "AdminPW"

#If you already have the tabcmd in your path variable, this should not be needed
#Otherwise, set this to the path where tabcmd is located
$env:Path = $env:Path + ";D:\Tableau\Tableau Server\9.3\bin"

#Make three directories to hold the various stages of the process
$downloadFolder = "C:\Automation\Download\"
$unzipFolder = "C:\Automation\Unzip\"
$completeFolder = "C:\Automation\Complete\"

#You will need both the short name and full name of the workbook
$workbookShortName = "SalesPerformance"
$workbookFullName = "Sales Performance"
$tableauLocation = "/workbooks/" + $workbookShortName + ".twb"
$localDownload = $downloadFolder + $workbookShortName + ".twb"
$unzippedWorkbook = $unzipFolder + $workbookFullName + ".twb"
$completeWorkbook = $completeFolder + $workbookFullName + ".twbx"



#Create directories
New-Item $downloadFolder -type directory -Force
New-Item $unzipFolder -type directory -Force
New-Item $completeFolder -type directory -Force

#Remove all existing objects in folder structure
Remove-Item -Path $downloadFolder'*' -Recurse -Force -Verbose
Remove-Item -Path $unzipFolder'*' -Recurse -Force -Verbose
Remove-Item -Path $completeFolder'*' -Recurse -Force -Verbose

tabcmd login -s http://$myTableauServer -u $myTableauServerUser -p $myTableauServerPW

tabcmd get $tableauLocation -f $localDownload

[System.IO.Compression.ZipFile]::ExtractToDirectory($localDownload, $unzipFolder)

#The parameter names can be discovered by opening the Tableau workbook in any text editor
$xmldata = New-Object XML
$xmldata.Load($unzippedWorkbook)
$parametersNode = $xmldata.workbook.datasources.datasource | where {$_.name -eq 'Parameters'}

#This sets the variables to the previous day and the first day of the month based on that previous day
$endDate = (get-date).AddDays(-1).ToString("#yyyy-MM-dd#")
$startDate = (Get-Date $endDate  -day 1 -hour 0 -minute 0 -second 0).ToString("#yyyy-MM-dd#")

$startDateNode = $parametersNode.column | where {$_.name -eq '[Parameter 1]'}
$startDateNode.value = $startDate
$endDateNode = $parametersNode.column | where {$_.name -eq '[Start Date (copy)]'}
$endDateNode.value = $endDate

$xmldata.Save($unzippedWorkbook)

[System.IO.Compression.ZipFile]::CreateFromDirectory($unzipFolder,$completeWorkbook)

tabcmd publish $completeWorkbook -n $workbookFullName -o -r "Default" --tabbed --db-username "dbUser" --db-password "dbPassword" --save-db-password

tabcmd logout
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
$workbookShortName = "Superstore"
$workbookFullName = "Superstore"
$tableauLocation = "/workbooks/" + $workbookShortName + ".twb"
$localDownload = $downloadFolder + $workbookShortName + ".twb"
$unzippedWorkbook = $unzipFolder + $workbookFullName + ".twb"
$completeWorkbook = $completeFolder + $workbookFullName + ".twbx"

New-Item $downloadFolder -type directory -Force
New-Item $unzipFolder -type directory -Force
New-Item $completeFolder -type directory -Force


#Remove all existing objects in folder structure
Remove-Item -Path $downloadFolder'*' -Recurse -Force -Verbose
Remove-Item -Path $unzipFolder'*' -Recurse -Force -Verbose
Remove-Item -Path $completeFolder'*' -Recurse -Force -Verbose

#Log into Tableau instance
tabcmd login -s http://$myTableauServer -u $myTableauServerUser -p $myTableauServerPW

tabcmd get $tableauLocation -f $localDownload

[System.IO.Compression.ZipFile]::ExtractToDirectory($localDownload, $unzipFolder)

#This section will need to be customized to the individual workbook based on the internal names of the parameters
#The parameter names can be discovered by opening the Tableau workbook in any text editor
$xmldata = New-Object XML
$xmldata.Load($unzippedWorkbook)
$parametersNode = $xmldata.workbook.datasources.datasource | where {$_.name -eq 'Parameters'}
$parametersNode = $parametersNode.column | where {$_.name -eq '[Base Salary]'}
$parametersNode.value = '75000'
$xmldata.Save($unzippedWorkbook)

[System.IO.Compression.ZipFile]::CreateFromDirectory($unzipFolder,$completeWorkbook)

tabcmd publish $completeWorkbook -n $workbookFullName -o -r "Tableau Samples" --tabbed --db-username "dbUser" --db-password "dbPassword" --save-db-password

tabcmd logout
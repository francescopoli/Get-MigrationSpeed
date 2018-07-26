--- current status: beta --- 
# Get-MigrationSpeed
Analyze a migration batch for speed on Exchange Online (analyze the initial full sync phase)

Usage:
clone or download the file.
dot source the script to import the function, like this:
. ./Get-MigrationSpeed

## Live analysis
Using the script against a live environment:
Connect to Exchange online PowerShell.

Run the analysis:<br>
<code>Get-MigrationSpeed</code>

If you want an expor of single entries to csv, use the parameters set:<br>
<code>Get-MigrationSpeed -ExportCSV -CSVFile ./filename.csv</code><br>
> (note -CSVFile parameter is optional, if omitted it will save to ./migration.csv)<br>
   
## Offline analysis
If you want to analyze an offline file:
Obtain the xml export of the batch to analyze from a live environment with:<br>
<code>Get-MoveRequest -batchname "migrationservice:batchname"| Get-MoveRequestStatistics | export-clixml c:\temp\migrationExport.xml</code><br>

Run the analysis against exported file:
<code>Get-MigrationSpeed -$XMLinput -$XmlFileName [path]\migrationExport.xml</code><br>

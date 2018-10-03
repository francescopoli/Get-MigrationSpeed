--- current status: beta --- 
# Get-MigrationSpeed
Analyze a migration batch for speed on Exchange Online (analyze the initial full sync phase)

Usage:<br>
Clone or download the file. <br>
dot source the script to import the function, like this:<br>
. ./Get-MigrationSpeed<br>

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

Run the analysis against exported file:<br>
<code>Get-MigrationSpeed -$XMLinput -$XmlFileName [path]\migrationExport.xml</code><br>


## Sample live execution:
<code>Get-MigrationSpeed</code>

```
_____Migration Batches_____
*                         *
    0: batch1
    1: batch2
    2: batch3
*_________________________*
Please make your selection: : 2


Batch selected: batch3
 
Server: serverName.contoso.com
____________________________________________________
Time taken(h): 0.113490900944444
Number of Mailboxes: 9
Total Mailboxes Size(Gb): 4.52
Total Data Transferred(Gb): 4.56
Move Quality: 99%
Move Performances (GB\h): 40.19
____________________________________________________
 
 
            Total migration performances
____________________________________________________
First Start: 06/27/2018 13:31:59
Last Complete: 06/27/2018 13:38:47
Time taken: 0d.0h:6m:48s
Number of Mailboxes: 9
Total Mailboxes Size(GB): 4.52
Total Data Transferred(GB): 4.56
Move Quality: 99%
Move Performances (GB\h): 40.19
____________________________________________________
```
> (Note: if multiple onPremises server are used, there will be an entry for each one, the Total migration performance is about the full batch itself, despite the used server.

## Versions

### BETA

* test in progress

### 1.0.0.0

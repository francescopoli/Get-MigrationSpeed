<#
    ###############Disclaimer#####################################################
    The sample scripts are not supported under any Microsoft standard support 
    program or service. The sample scripts are provided AS IS without warranty  
    of any kind. Microsoft further disclaims all implied warranties including,  
    without limitation, any implied warranties of merchantability or of fitness for 
    a particular purpose. The entire risk arising out of the use or performance of  
    the sample scripts and documentation remains with you. In no event shall 
    Microsoft, its authors, or anyone else involved in the creation, production, or 
    delivery of the scripts be liable for any damages whatsoever (including, 
    without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use 
    of or inability to use the sample scripts or documentation, even if Microsoft 
    has been advised of the possibility of such damages.
    ###############Disclaimer#####################################################
#>
<#
	Script is in beta
#>
<#       
    .SYNOPSIS
    Calculate the migration troughput during an Exchange hybrid migration, based on the 
    initial FullSync.

    .DESCRIPTION
    Used to calculate the throughput of the initial FullSync operation for a migration
    batch, on an Exchange hybrid migration.
    Provide the calculation for the batch performance and for each MRSserver used onPremises

    Allow to import an exported diagnostic for later processing
    
    Usage
    load the function in PowerShell using
    . .\Get-MigrationSpeed.ps1

    then use the imported function using:
    Get-MigrationSpeed [parameters]

	Author: Francesco Poli fpoli@microsoft.com
	
    .PARAMETER ExportCSV
    [ParameterSet] Export
    If switch is passed in, the script expect a valid path and filename in CSVFile parameter
    
    .PARAMETER CSVFile
    [ParameterSet] Export
    A valid path and file filename where to save migration speed data.
    If omitted, the file will be generated in the execution folder, with name ./migration.csv

    .PARAMETER XMLInput
    [ParameterSet] InputFile
    If switch is passed in, the script expect a valid XML file in XmlFileName parameter
      
    .PARAMETER XmlFileName
    [ParameterSet] InputFile
    A valid export of a migration diagnostics taken with
    Get-MoveRequest -BatchName "migrationservice:BatchName | Get-MoveRequestStatistics | Export-Clixml [path]\file.xml

    .EXAMPLE
    Get-MigrationSpeed

    .EXAMPLE
    Get-MigrationSpeed -ExportCSV -CSVFile c:\temp\export.csv

    .EXAMPLE
    Get-MigrationSpeed -XMLInput -XmlFileName c:\temp\data.xml 
    
    .EXAMPLE
    Get-MigrationSpeed -ExportCSV -CSVFile c:\temp\export.csv -XMLInput -XmlFileName c:\temp\data.xml 

#>

function Get-MigrationSpeed {
    [cmdletbinding(
        DefaultParameterSetName='liveNoexport'
    )]

    param(
        [Parameter(ParameterSetName='liveNoexport',Mandatory=$false)]
            [switch] $LiveNoExport = $true,
        [Parameter(ParameterSetName='Export',Mandatory=$false)]
            [switch]$ExportCSV,
        [Parameter(ParameterSetName='Export',Mandatory=$false)]
            [string]$CSVFile = ".\migration.csv",
        [Parameter(ParameterSetName='InputFile',Mandatory=$false)]
            [switch]$XMLInput,
        [Parameter(ParameterSetName='InputFile',Mandatory=$false)]
            [string]$XmlFileName
    )

    function EBTo-Number{
	     param(
            $exoBytes
         )
         $fix = $exoBytes.ToString().split("(")[1].split(" ")[0].replace(",","");
	     return [long]$fix/1024/1024;
    }

    if (!$XMLInput){
        $x=0;
        Write-Host "                           "
        Write-Host "_____Migration Batches_____"
        Write-Host "*                         *"
        $batches = Get-MigrationBatch
        foreach($batch in $batches){
            Write-host ("    {0}: {1}" -f $x,$batch.identity)
            $x++
           }
        Write-Host "*_________________________*"

        $batchNumber = 0
        $reAsk = $true
        While ($reAsk){
            $batchNumber = Read-Host "Please make your selection: "
            if ( ( $batchNumber -gt $batches.Count-1) -or 
                 ( $batchNumber -lt 0) ) {
                Write-Host "Not a valid batch number, selection must be between 0 and $($batches.count)"
            }
            else{
                $reAsk = $false
                Write-Host "Batch selected: $($batches[$batchNumber].identity)" -ForegroundColor Yellow
            }
        }

        $migrations = Get-MoveRequest -BatchName "migrationservice:$($batches[$batchNumber].identity)" |
                        Get-MoveRequestStatistics
    }
    else {
        if (Test-Path $XmlFileName) {
            $migrations = Import-Clixml $XmlFileName
        }
        else{
            Write-Host "XML file seems to not be valid, please check."
            throw "XML file seems to not be valid, please check.";
        }
    }

    if ($ExportCSV) {
        if ( !(Test-Path $CSVFile.Substring(0,$CSVFile.LastIndexOf("\") ) ) ){
            Write-Host "Export csv path is invalid, please verify it"
            Throw "Export csv path is invalid, please verify it"
        }
    }

    $SourceServers = $migrations | Select-Object SourceServer -Unique
    $glob_InitialSeedingCompletedTimestamp = 0
    $glob_StartTimestamp = 0
    $glob_totalMbxSize = 0
    $glob_totalDataTransferred = 0
    $glob_totalStalledTime = 0
    
    $obj = @()
    foreach($souS in $SourceServers){
        $migs = $migrations | where{$_.sourceServer -eq $souS.sourceserver}
        $StartTimestamp = 0
        $InitialSeedingCompletedTimestamp = 0
        $totalMbxSize = 0
        $totalDataTransferred = 0
        $totalStalledTime


        foreach($mig in $migs){
            #global start stop seeding
            if ( $glob_StartTimestamp -ne 0 ){ 
                if($mig.StartTimestamp -lt $glob_StartTimestamp ){
                        $glob_StartTimestamp = $mig.StartTimestamp
                }  
            }
            else{ 
                $glob_StartTimestamp = $mig.StartTimestamp 
            }

            if ( $mig.InitialSeedingCompletedTimestamp -gt $glob_InitialSeedingCompletedTimestamp ) {
                $glob_InitialSeedingCompletedTimestamp = $mig.InitialSeedingCompletedTimestamp
            }

            #local per server start stop seeding
            if ( $StartTimestamp -ne 0 ){ 
                if($mig.StartTimestamp -lt $StartTimestamp ){
                        $StartTimestamp = $mig.StartTimestamp
                }  
            }
            else{ 
                $StartTimestamp = $mig.StartTimestamp 
            }

            if ( $mig.InitialSeedingCompletedTimestamp -gt $InitialSeedingCompletedTimestamp ) {
                $InitialSeedingCompletedTimestamp = $mig.InitialSeedingCompletedTimestamp
            }

            $MbxSize = [Math]::Round((EBTo-Number $mig.TotalMailboxSize));
            if ($MbxSize -eq 0) {$MbxSize=1}
            $totalMbxSize += $MbxSize
            $glob_totalMbxSize += $MbxSize


            $DataTransferred = [Math]::Round((EBTo-Number $mig.BytesTransferred)) #this is in MB
            $totalDataTransferred += $DataTransferred
            $glob_totalDataTransferred += $DataTransferred
            
            $TotalStalledTime += 
                $mig.TotalStalledDueToContentIndexingDuration.TotalMinutes + 
                $mig.TotalStalledDueToMailboxLockedDuration.TotalMinutes + 
                $mig.TotalStalledDueToReadThrottle.TotalMinutes +
                $mig.TotalStalledDueToWriteThrottle.TotalMinutes + 
                $mig.TotalStalledDueToReadCpu.TotalMinutes + 
                $mig.TotalStalledDueToWriteCpu.TotalMinutes + 
                $mig.TotalStalledDueToReadUnknown.TotalMinutes + 
                $mig.TotalStalledDueToWriteUnknown.TotalMinutes;
            
            $glob_totalStalledTime += $TotalStalledTime

            if ($DataTransferred -eq 0) { $DataTransferred = 1}
            $matrix = [ordered]@{
                "user" = $mig.alias;
                "sourceServer" = $mig.sourceServer;
                "TimeInMin"= $mig.TotalInProgressDuration.TotalMinutes;
                "MbxSize" = $MbxSize;
                "DataTransferred" = $DataTransferred; 
                "MovePerfDataPerMinutes" = $DataTransferred / $mig.TotalInProgressDuration.TotalMinutes;
                "MoveQuality" = $MbxSize/$DataTransferred;
                "RemoteDatabaseName" = $mig.RemoteDatabaseName;
                "TargetDatabase" = $mig.TargetDatabase;
                "ItemsTransferred" = $mig.ItemsTransferred;
                "TotalStalledTime" = $TotalStalledTime;
                "TotalTransientFailureDuration" = $mig.TotalTransientFailureDuration.TotalMinutes;
                "TimeToMove" = ($mig.InitialSeedingCompletedTimestamp - $mig.StartTimestamp).TotalMinutes;
                "MBPerMinutes" = ($DataTransferred) / 
                                 ($mig.InitialSeedingCompletedTimestamp - $mig.StartTimestamp).TotalMinutes;
            }

            if ($ExportCSV) {
                $obj += New-Object -TypeName PSObject -Property $matrix
            }
        }
       
    
        Write-Host " "
        Write-Host "Server: $($souS.sourceserver)" 
        Write-Host "____________________________________________________"
        Write-Host "Time taken(h): $((($InitialSeedingCompletedTimestamp - $StartTimestamp).totalhours))"
        Write-Host "Number of Mailboxes: $($migs.count)" 
        Write-Host "Total Mailboxes Size(Gb): $([Math]::round($totalMbxSize/1024,2))"
        Write-Host "Total Data Transferred(Gb): $([Math]::round($totalDataTransferred/1024,2))"
        Write-Host "Move Quality: $([Math]::round($totalMbxSize/$totalDataTransferred,2)*100)%"
        Write-Host "Move Performances (GB\h): $( [Math]::round((($totalDataTransferred/1024) / 
                                                               ($InitialSeedingCompletedTimestamp - $StartTimestamp).Totalhours),2) )" 
        Write-Host "____________________________________________________"
        Write-Host " "
    }
        
        $tt = $glob_InitialSeedingCompletedTimestamp - $glob_StartTimestamp
    # global results
    Write-Host " "
    Write-Host "            Total migration performances" 
    Write-Host "____________________________________________________"
    Write-Host "First Start: $($glob_StartTimestamp)" 
    Write-Host "Last Complete: $($glob_InitialSeedingCompletedTimestamp)"
    Write-Host ("Time taken: {0}d.{1}h:{2}m:{3}s" -f $tt.days,$tt.hours,$tt.minutes,$tt.seconds )
    Write-Host "Number of Mailboxes: $($migrations.count)" 
    Write-Host "Total Mailboxes Size(GB): $([Math]::round($glob_totalMbxSize/1024,2))"
    Write-Host "Total Data Transferred(GB): $([Math]::round($glob_totalDataTransferred/1024,2))"
    Write-Host "Move Quality: $([Math]::round($glob_totalMbxSize/$glob_totalDataTransferred,2)*100)%"
    Write-Host "Move Performances (GB\h): $([Math]::round( ($glob_totalDataTransferred/1024) /
                                                           ((($glob_InitialSeedingCompletedTimestamp - $glob_StartTimestamp).Totalminutes)/60),2) )" 
    Write-Host "____________________________________________________"

    if ($ExportCSV){
        
         $obj | Export-Csv -NoTypeInformatio -Path $CSVFile
    }
}

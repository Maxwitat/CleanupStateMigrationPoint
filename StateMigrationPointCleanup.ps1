#####================================================================================
## Title: StatMigrationPointCleanup.ps1
##    
##    **********************************************************************************************************
##       This script is provided as is with no warrenties. Please test carefully before applying it to a 
##       a productive environment
##
##    ***********************************************************************************************************
##
##    Author: Frank Maxwitat
##    Version : 1.0,  11.06.2020
##    
##    Checks SMP for orphaned folders and creates a cleanup file
#>

#--------------------Define variables district -------------------------------------------------

$WhatIfPreference = $true     #No deletion takes place unless you set it to false! Test thoroughly
$DBServerName = 'SVRP01' 
$DBInstance = '' #Leave open if default
$DBName = 'CM_P01'
$USMTPath = '\\SVRCMMulti\c$\USMT'
$CheckFoldersOlderThan = '-1' #in days
$Global:MaxLogSizeInKB = 250 
$Global:LogFile = $PSScriptRoot + '\StatMigrationPointCleanup.log'

#--------------------End variables district-----------------------------------------------------

#--------------------Functions -----------------------------------------------------------------

function WriteLog
{
  param (
  [Parameter(Mandatory=$true)]
  $message,
  [Parameter(Mandatory=$true)]
  $component,
  [Parameter(Mandatory=$true)]
  $type )
  switch ($type)
  {
    1 { $type = "Info" }
    2 { $type = "Warning" }
    3 { $type = "Error" }
    4 { $type = "Verbose" }
  }
  if (($type -eq "Verbose") -and ($Global:Verbose))
  {
    $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($Global:ScriptName + ":" + $component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
    $toLog | Out-File -Append -Encoding UTF8 -FilePath ("filesystem::{0}" -f $Global:LogFile)
    Write-Host $message
  }
  elseif ($type -ne "Verbose")
  {
    $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($Global:ScriptName + ":" + $component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
    $toLog | Out-File -Append -Encoding UTF8 -FilePath ("filesystem::{0}" -f $Global:LogFile)
    Write-Host $message
  }
  if (($type -eq 'Warning') -and ($Global:ScriptStatus -ne 'Error')) { $Global:ScriptStatus = $type }
  if ($type -eq 'Error') { $Global:ScriptStatus = $type }
  if ((Get-Item $Global:LogFile).Length/1KB -gt $Global:MaxLogSizeInKB)
  {
    $log = $Global:LogFile
    Remove-Item ($log.Replace(".log", ".lo_"))
    Rename-Item $Global:LogFile ($log.Replace(".log", ".lo_")) -Force
  }
} 

#--------------------End Functions -------------------------------------------------------------

if(Test-Path $PSScriptRoot\$OutFileName)
{
    Remove-Item -Path $PSScriptRoot\$OutFileName
}

if($DBInstance -ne '')
{
    $DBConnect = $DBServerName + '\' + $DBInstance
}
else
{
    $DBConnect = $DBServerName
}


$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset
$con = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=$DBName;Data Source=$DBConnect"

$Folders = Get-ChildItem -Path $USMTPath -Directory | where {$_.LastWriteTime -le $(get-date).Adddays($CheckFoldersOlderThan)} 

Write-Host 'Number of USMT folders: '  $Folders.Count  -BackgroundColor DarkGreen -ForegroundColor White
WriteLog -message 'Number of USMT folders: '  $Folders.Count -type Info

$objConnection.Open($con)
$objConnection.CommandTimeout = 0
# *********** Check If connection is open *******************
If($objConnection.state -eq 0)
{
	Write-Host "Error: SCCM Central DB ServerName or Central SCCM DB Name is not properly set or your account does not have sufficient access"
	Add-Content $logfile -Value "Error: Central SCCM DB ServerName or Central SCCM DB Name is not properly mentioned in Config XML File or Your Account does not have sufficient Access"
	Exit 1        
}
else
{
    foreach($folder in $folders)
    {
        Write-Host 'Checking folder ' $folder.Name -BackgroundColor DarkGreen -ForegroundColor White
$strSQL = @"
    select StorePath, StoreSize, MigrationStatus, SourceName from v_StateMigration where StorePath like '%$folder%'
"@

        Try
        {
            $objRecordset.Open($strSQL,$objConnection)
		    $objRecordset.MoveFirst()
		    $rows=$objRecordset.RecordCount

            $StorePath = $objRecordset.Fields.Item(0).Value
            $StoreSize = $objRecordset.Fields.Item(1).Value
            $MigrationStatus = $objRecordset.Fields.Item(2).Value
            $SourceName = $objRecordset.Fields.Item(3).Value
        
            Write-Host "Found object in database. Path " $StorePath "; Size " $StoreSize "; MigrationStatus " $MigrationStatus "; SourceName " $SourceName -BackgroundColor DarkGreen -ForegroundColor White

            $objRecordset.Close()
        }
        Catch
        {
            WriteLog -message 'Orphaned folder: ' $USMTPath\$folder -type Info
            Remove-Item -Path $USMTPath\$folder -Recurse
            Write-Host "Orphaned folder: " $folder -BackgroundColor black -ForegroundColor White
        }
    }
}

$objConnection.Close()



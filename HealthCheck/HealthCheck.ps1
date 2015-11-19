##################################################################
#
#  Name:      HealthCheck.ps1
#  Author:    Omid Afzalalghom
#  Date:      02/09/2015
#  Requires:  Excel, PS, SQL Tools.
#  Revisions: 
##################################################################

#Requires –Version 2.0
[CmdletBinding()]
Param(
    [Parameter(Position=0,Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias('Instance')]
    [string] $SQLInstance,

    [Parameter(Position=1,Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias('DBName')]
    [string] $DatabaseName
)

# Fail if sqlps module cannot be found (allows use of invoke-sqlcmd)
try   { Import-Module “sqlps” -DisableNameChecking -ErrorAction stop }
catch { throw "Exiting because the sqlps module cannot be found" }
   
#Assign variable values.
$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path    
$dir = Join-Path $PSScriptRoot "\"
$date = Get-Date -Format "yyyyMMdd"


TRY
{
    $ErrorActionPreference = "Stop"

    switch ($SQLInstance)
    {
        #Replace backslash in instance name.
        {$_.IndexOf("\") -gt 0} { $SQLInstance = $SQLInstance.replace('\','$') }

        #Replace comma if present due to port number being included.
        {$_.IndexOf(",") -gt 0} { $SQLInstance = $SQLInstance.replace(',','$') }

        #No replacements required
        Default { Write-Verbose "No transformations on instance name necessary" }
    }

    $Report = "$dir\Reports\Report_${SQLInstance}_${DatabaseName}_$date.xlsm"

    Write-Verbose 'Checking database is online.'
    
    $dbState = invoke-sqlcmd –ServerInstance $SQLInstance -Query "SELECT state_desc FROM sys.databases WHERE name = '$DatabaseName';"  | %{'{0}' -f $_[0]}
    if ($dbState -ne "online")
    {
        "The database '$DatabaseName' is in $dbState mode. Please bring database online and try again."
        sleep -s 5
        return
    }

    #Get SQL Server version and compatibility level of target database.
    [int]$compLevel = invoke-sqlcmd –ServerInstance $SQLInstance -Query "SELECT compatibility_level FROM sys.databases WHERE name = '$DatabaseName';"  | %{'{0}' -f $_[0]}
    [int]$ver = invoke-sqlcmd –ServerInstance $SQLInstance -Query "SELECT REPLACE(LEFT(CONVERT(varchar, SERVERPROPERTY ('ProductVersion')),2), '.', '');"  | %{'{0}' -f $_[0]}

    #Set folder path depending upon SQL version.
    switch ($ver)
    {
        9  { $FolderPath = "$dir\Scripts_2005\"; break }
        10 { $FolderPath = "$dir\Scripts_2008\"; break }
        11 { $FolderPath = "$dir\Scripts_2012\"; break }
        12 { $FolderPath = "$dir\Scripts_2014\"; break }
        Default { throw "Unknown SQL Server version. Exiting" }
    }

    #Run server information script.
    Write-Verbose 'Getting server information.'
    powershell.exe -file "$FolderPath\ServerInformation.ps1" $SQLInstance $PSScriptRoot

    $Files = Get-ChildItem  -Path $FolderPath -Filter "*.sql" | Sort-Object
    $Count = $Files.length

    #Loop through the .sql files and run them.
    foreach ($Filename in $Files)
    { 
        $i++
        $Outfile = Join-Path "$dir\Reports\Temp\"  $Filename.Name.Replace(".sql", ".csv")

        invoke-sqlcmd –ServerInstance $SQLInstance -Database $DatabaseName -InputFile $Filename.Fullname | Export-Csv -Path $Outfile -NoTypeInformation -Encoding UTF8
        Write-Progress -activity "Exporting $Filename to CSV. File $i/$Count." -status "Completed: " -PercentComplete (($i/$Count)*100)       
    }

    #For SQL 2005/2008 compatibility 80 databases, run scripts that use CROSS APPLY or UNPIVOT against the master database.
    if ( ($ver -eq 9) -OR ($ver -eq 10) )
    {
        $FolderPath = "$FolderPath\80\"
        $Files = Get-ChildItem -Path $FolderPath -Filter "*.sql" | Sort-Object

        foreach ($Filename in $Files)
        {
            $Outfile = Join-Path "$dir\Reports\Temp\"  $Filename.Name.Replace(".sql", ".csv")
            invoke-sqlcmd –ServerInstance $SQLInstance -Database "master" -InputFile $Filename.Fullname -Variable "DB='${DatabaseName}'" | Export-Csv -Path $Outfile -NoTypeInformation -Encoding UTF8   
        }
    }

    #Create report file.
    Copy-Item "$dir\Reports\Master.xlsm" -Destination $Report -Force 

    #Open report file.
    $excel = New-Object -ComObject excel.application
    $excel.Visible = $False
    $workbook = $excel.workbooks.open($report)   
    $worksheet = $workbook.worksheets.item(1)

    #Run import macro.
    $excel.Run("loader")
     
    #Run formatting macro.
    $excel.Run("FormatWorksheets")

    #Clean up.
    $workbook.close()
    $excel.quit()
    $workbook = $Null
    $worksheet = $Null
    $excel = $Null
     
    #Remove CSVs.
    Get-ChildItem -Path "$dir\Reports\Temp\" -Recurse -Filter "*.csv" | Remove-Item
    
}
CATCH 
{
    Get-ChildItem -Path "$dir\Reports\Temp\" -Recurse -Filter "*.csv" | Remove-Item
    Write-output "$(Get-Date –f G) - Error:$Filename - $($_.Exception.Message)" | Out-File -FilePath "$dir\ErrorLog.txt" -Append
}

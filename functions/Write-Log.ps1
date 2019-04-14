Function Write-Log
{
    param(
    [Parameter(Mandatory = $true)]
    [string]$LogString,
    [Parameter(Mandatory = $false)]
    [ValidateSet("Emergency", "Alert", "Critical", "Error", "Warning", "Notice", "Informational", "Debug")]
    [string]$Severity = "Informational"
    )
    
    #$scriptLogFile = $env:APPDATA + "\user_mgmt.log.txt"
    #$sysLogServer = $null
    #$maxLogFileSize = 10000000

    # Rotate local log file    
    #if((Test-Path $scriptLogFile) -and ((Get-Item $scriptLogFile).Length -gt $maxLogFileSize))
    #{        
    #    $archiveFile = $scriptLogFile + ".1"
    #    if((Test-Path $archiveFile))
    #    {            
    #        Remove-Item -Path $archiveFile
    #    }        
    #    Move-Item -Path $scriptLogFile -Destination $archiveFile        
    #}
        
    [int]$intSeverity = 0
    
    if($Severity -eq "Emergency") { $intSeverity = 0; $color = "Red" }
    if($Severity -eq "Alert") { $intSeverity = 1; $color = "Red" }
    if($Severity -eq "Critical") { $intSeverity = 2; $color = "Red" }
    if($Severity -eq "Error") { $intSeverity = 3; $color = "Red" }
    if($Severity -eq "Warning") { $intSeverity = 4; $color = "Magenta" }
    if($Severity -eq "Notice") { $intSeverity = 5; $color = "Cyan" }
    if($Severity -eq "Informational") { $intSeverity = 6; $color = "White" }
    if($Severity -eq "Debug") { $intSeverity = 7; $color = "Yellow" }       
    
    $Facility = 22
    $Priority = ([int]$Facility * 8) + [int]$intSeverity
                
    $TimeStamp = ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
    $LogString = ($TimeStamp + " [" + $Severity + "] [" + $((Get-PSCallStack)[1].Command) + "] " + ($LogString -replace "`n"," "))    
    
    if($intSeverity -le 5)
    {
        Write-Host $LogString -Foreground $color       
    }
    elseif(($intSeverity -eq 6) -and ((!(Test-Path ("variable:script:settings"))) -or ($script:settings.ShowVerboseOutput)))
    {
        Write-Host $LogString -Foreground $color
    }
    elseif(($intSeverity -eq 7) -and ((!(Test-Path ("variable:script:settings"))) -or ($script:settings.ShowDebugOutput)))
    {
        Write-Host $LogString -Foreground $color       
    }        
}
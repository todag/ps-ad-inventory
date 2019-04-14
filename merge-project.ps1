###############################################################################
#.SYNOPSIS
#   Merges the project into one single .ps1 file.
#   Offers option to sign the resulting file if a code signing
#   certificate is found.
#
#   This is just a quick hack...
#
###############################################################################

#$VerbosePreference = "Continue"

#
# The merged project till be written to this file
#
$outputFile = "ad-inventory-merged.ps1"


$data = (Get-Content ".\ad-inventory.ps1")
[System.Collections.Generic.List[string]]$script:outputData = New-Object System.Collections.Generic.List[string]


#region Functions
function Add-Functions
{
    Write-Output "Merging functions"
    $script:outputData.Add("#region Functions")
    Get-ChildItem -Path .\functions -Filter *.ps1 | ForEach-Object {
        Write-Verbose ("Adding function: " + $_.Name)
        $function = (Get-Content $_.FullName)
        foreach($line in $function)
        {
            $script:outputData.Add($line)
        }
    }
    $script:outputData.Add("#endregion")
}

function Add-Classes
{
    Write-Output "Merging classes"
    $script:outputData.Add("#region Classes")
    Get-ChildItem -Path .\classes -Filter *.ps1 | ForEach-Object {
        Write-Verbose ("Adding class: " + $_.Name)
        $function = (Get-Content $_.FullName)
        foreach($line in $function)
        {
            $script:outputData.Add($line)
        }
    }
    $script:outputData.Add("#endregion")
}

function Add-XAML
{
    Write-Output "Merging XAML resources"
    $script:outputData.Add("#region XAML")
    Write-Verbose "Adding MainWindow.xaml"
    $script:outputData.Add("[xml]`$script:xamlMainWindow = @`"" )
    foreach($line in (Get-Content -Path .\resources\MainWindow.xaml))
    {
        $script:outputData.Add($line)
    }

    $script:outputData.Add("`"@")

    Write-Verbose "Adding SettingsWindow.xaml"
    $script:outputData.Add("[xml]`$script:xamlSettingsWindow = @`"" )
    foreach($line in (Get-Content -Path .\resources\SettingsWindow.xaml))
    {
        $script:outputData.Add($line)
    }
    $script:outputData.Add("`"@")

    Write-Verbose "Adding ExportWindow.xaml"
    $script:outputData.Add("[xml]`$script:xamlExportWindow = @`"" )
    foreach($line in (Get-Content -Path .\resources\ExportWindow.xaml))
    {
        $script:outputData.Add($line)
    }
    $script:outputData.Add("`"@")

    Write-Verbose "Adding LapsWindow.xaml"
    $script:outputData.Add("[xml]`$script:xamllapsWindow = @`"" )
    foreach($line in (Get-Content -Path .\resources\LapsWindow.xaml))
    {
        $script:outputData.Add($line)
    }
    $script:outputData.Add("`"@")

    Write-Verbose "Adding DebugWindow.xaml"
    $script:outputData.Add("[xml]`$script:xamldebugWindow = @`"" )
    foreach($line in (Get-Content -Path .\resources\DebugWindow.xaml))
    {
        $script:outputData.Add($line)
    }
    $script:outputData.Add("`"@")    
    $script:outputData.Add("#endregion")
}

function Add-CSharp
{
    Write-Output "Merging CSharp resources"
    $script:outputData.Add("#region CSharp")
    $script:outputData.Add("`$script:converters = @`"" )
    foreach($line in (Get-Content -Path .\resources\IValueConverters.cs))
    {
        $script:outputData.Add($line)
    }
    $script:outputData.Add("`"@")
    $script:outputData.Add("#endregion")
}
#endregion

[bool]$skip = $false
[string]$skipUntil = $null
foreach($line in $data)
{
    if($skip -eq $true -and $line.StartsWith($skipUntil))
    {
        $skip = $false
        $skipUntil = $null
        continue
    }
    elseif($skip -eq $true -and !$line.StartsWith($skipUntil))
    {
        continue
    }

    if($line.StartsWith("#region DotSource"))
    {
        $skip = $true
        $skipUntil = "#endregion"
        Add-Functions
    }
    elseif($line.StartsWith("#region Classes"))
    {
        $skip = $true
        $skipUntil = "#endregion"
        Add-Classes
    }
    elseif($line.StartsWith("#region XAML"))
    {
        $skip = $true
        $skipUntil = "#endregion"
        Add-XAML
    }
    elseif($line.StartsWith("#region CSharp"))
    {
        $skip = $true
        $skipUntil = "#endregion"
        Add-CSharp
    }
    else
    {
        $script:outputData.Add($line) | Out-Null
    }
}

#
# Write $script:outputData to $outputFile
#
$script:outputData.Add(("#").PadRight(70,"#") + "#")
$script:outputData.Add(("#").PadRight(70," ") + "#")
$script:outputData.Add(("# Merged by user: " + $env:USERNAME).PadRight(70," ") + "#")
$script:outputData.Add(("# On computer:    " + $env:COMPUTERNAME).PadRight(70," ") + "#")
$script:outputData.Add(("# Date:           " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")).PadRight(70," ") + "#")
if((Get-ChildItem cert:\CurrentUser\My -codesign).Count -eq 0)
{
    $script:outputData.Add(("# No code signing certificate found!").PadRight(70," ") + "#")
}

# Just for fun, calculate real code lines, exluding empty lines and comments
$count = 0
foreach($row in $script:outputData)
{
    if(!($row.StartsWith("#")) -and !([string]::IsNullOrWhiteSpace($row)))
    {
       $count++
    }
}
$script:outputData.Add(("# LoC: " + $script:outPutData.Count.ToString() + " of which " + $count.ToString() + " is not comments or whitespace").PadRight(70," ") + "#")
$script:outputData.Add(("#").PadRight(70," ") + "#")
$script:outputData.Add(("#").PadRight(70,"#") + "#")

Set-Content $outputFile $script:outputData
Write-Output ("Merged project file created")
Write-Output ("Total line count: " + $script:outputData.Count)
Write-Output ("Code line count: " + $count.ToString())

#
# Check if code signing certificate exists and ask if merged script should be signed
#
if((Get-ChildItem cert:\CurrentUser\My -codesign).Count -gt 0)
{
    $answer = Read-Host "Found code signing certificate, sign merged file? (y/n)"
    if($answer -eq "y")
    {
        $index = 0
        foreach($cert in (Get-ChildItem cert:\CurrentUser\My -codesign))
        {
            Write-Output ("Index [" + $index.ToString() + "] Subject: " + $cert.Subject + " TP: " + $cert.Thumbprint)
        }

        [int]$certIndex = Read-Host "Type index of certificate to sign with: "
        Set-AuthenticodeSignature $outputFile @(Get-ChildItem cert:\CurrentUser\My -codesign)[$certIndex]
    }
}
else
{
    Write-Output "*** No code signing certificate found! ***"
}
Write-Output ("Merged file size: " + (Get-Item $outputFile).Length + " bytes")
Write-Output "Merge operations finished. Script terminated, press enter to exit session..."
Read-Host
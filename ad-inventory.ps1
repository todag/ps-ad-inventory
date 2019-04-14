# MIT License
#
# Copyright (c) 2019 https://github.com/todag

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

###############################################################################
#.SYNOPSIS
#
#    This application allows you to easily visualise and edit Active Directory
#    Computer and User attributes. The logic is mainly written in Powershell
#    and the UI is defined in XAML/WPF. Some hints of C# (mainly for IValueConverters)
#
#.DESCRIPTION
#
#.NOTES
#
#    Version:        0.6
#
#    Author:         <https://github.com/todag>
#
#    Creation Date:  <2019-04-14>
#              0.5:  <2017-12-27>
#
###############################################################################

Set-StrictMode -Version 2.0
$script:appVersion = "AD Inventory - V0.6 (c) 2019 todag"
Write-Host $script:appVersion
#region DotSource (Don't remove this region, used by merge-project.ps1)
Write-Host "Dot sourcing functions in .\functions..."

#
# DotSource all files in .\functions
#
Get-ChildItem -Path .\functions -Filter *.ps1 | ForEach-Object {
    Write-Host ("Dot sourcing function: " + $_.Name)
    . $_.FullName
}
Write-Log -LogString "Dot sourcing classes in .\classes..." -Severity "Debug"

#
# DotSource all files in .\classes
#
Get-ChildItem -Path .\classes -Filter *.ps1 | ForEach-Object {
    Write-Log -LogString ("Dot sourcing class: " + $_.Name) -Severity "Debug"
    . $_.FullName
}
#endregion

#region Classes (Don't remove this region, used by merge-project.ps1)
#endregion

#region XAML Load XAML Resources (Don't remove this region, used by merge-project.ps1)
[xml]$script:xamlMainWindow = Get-Content -Path .\resources\MainWindow.xaml
[xml]$script:xamlSettingsWindow = Get-Content -Path .\resources\SettingsWindow.xaml
[xml]$script:xamlExportWindow = Get-Content -Path .\resources\ExportWindow.xaml
[xml]$script:xamlLapsWindow = Get-Content -Path .\resources\LapsWindow.xaml
[xml]$script:xamlDebugWindow = Get-Content -Path .\resources\DebugWindow.xaml
#endregion

#region CSharp (Don't remove this region, used by merge-project.ps1)
[string]$script:converters = Get-Content -Path .\resources\IValueConverters.cs
#endregion

#
# ---------------------- Script scope variables ----------------------
#
$script:dateTimeFormat = "yyyy-MM-dd HH:mm"
$script:schema = $null
$script:MainWindow = @{}

#
# IValueConverters
#
$script:FileTimeConverter = $null
$script:DateFormatConverter = $null
$script:ADPropertyValueCollectionConverter = $null
$script:msExchRemoteRecipientTypeConverter = $null
$script:msExchRecipientDisplayTypeConverter = $null
$script:managedByConverter = $null
$script:settingsFile = ($env:appdata + "\PS-AD-Inventory\settings.xml")

[Settings]$script:settings = Get-Settings
# --------------------------------------------------------------------

#
# Check if ActiveDirectory module is installed, if not, exit script
#
Write-Verbose "Importing Active Directory Powershell Module..."
Import-Module ActiveDirectory -Verbose:$false -ErrorAction SilentlyContinue | Out-Null
if(!$?)
{
    Write-Log -LogString "Failed to load ActiveDirectory module! Script will terminate." -Severity "Critical"
    Read-Host
    Exit
}

Write-Log -LogString "Loading assemblies and converters..." -Severity "Notice"
try
{
    Add-Type -AssemblyName PresentationFramework

    # This is required to be able to show/hide console window.
    Add-Type -TypeDefinition $script:converters -ReferencedAssemblies PresentationFramework, Microsoft.ActiveDirectory.Management -Verbose
    Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
}
catch
{
    Write-Log -LogString ($_.Exception.Message) -Severity "Critical"
    Write-Log -LogString "Press enter to exit..." -Severity "Notice"
    Read-Host
    Exit
}

Show-MainWindow
Write-Host "**** Script has terminated... ****"

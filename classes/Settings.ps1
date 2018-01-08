###############################################################################
#
#.SYNOPSIS
#    This class holds various application settings, including Attribute Definitions
#
###############################################################################

Class Settings
{
    Settings()
    {
        $this.ComputerAttributeDefinitions = New-Object System.Collections.Generic.List[AttributeDefinition]
        $this.UserAttributeDefinitions = New-Object System.Collections.Generic.List[AttributeDefinition]
    }
    [System.Collections.Generic.List[AttributeDefinition]] $ComputerAttributeDefinitions
    [System.Collections.Generic.List[AttributeDefinition]] $UserAttributeDefinitions
    [bool] $OnStartLoadGroups = $true
    [bool] $OnStartLoadOrganizationalUnits = $true
    [bool] $ShowVerboseOutput = $false
    [bool] $ShowDebugOutput = $false
    [int] $ComputerInactiveLimit = 90
    [int] $UserInactiveLimit = 90
}
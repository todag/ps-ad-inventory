###############################################################################
#
#.SYNOPSIS
#    This class defines an AD Attribute.
#
#.PARAMETER FriendlyName
#    Friendly name of attribute, will be shown as header in UI and exports..
#
#.PARAMETER Attribute
#    Name of AD Attribute.
#
#.PARAMETER IsEditable
#    Whether or not editing should be allowed.
#
#.PARAMETER DisplayIn
#    Whether to display this attribute in DataGrid och Sidebar.
#
###############################################################################

Class AttributeDefinition
{
    AttributeDefinition([string]$FriendlyName, [string]$Attribute, [bool]$IsEditable, [string]$DisplayIn)
    {
       $this.FriendlyName = $FriendlyName
       $this.Attribute = $Attribute
       $this.IsEditable = $IsEditable
       $this.DisplayIn = $DisplayIn
    }

    AttributeDefinition()
    {
        $this.IsEditable = $false
        $this.DisplayIn = "DataGrid"
    }

    [string] $FriendlyName
    [string] $Attribute
    [bool] $IsEditable
    [string] $DisplayIn
    [bool] $IgnoreConverter

    #These properties will be loaded from Active Directory Schema
    [string]$Syntax
    [bool]$IsSingleValued
}

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
#region Functions
###############################################################################
#.SYNOPSIS
#    Returns an IValueConverter for an attribute.
#.DESCRIPTION
#    Currently contains converters for:
#    DateFormatConverter (for attributes with Syntax GeneralizedTime)
#    FileTimeConverter (for attributes that stores Date as FileTime)
#    ADPropertyValueCollectionConverter (For Multivalue attributes)
#    msExchRemoteRecipientTypeConverter (For attribute msExchRemoteRecipientTypeConverter)
#
#.PARAMETER AttributeDefinition
#    AttributeDefinition of the attribute to get the converter for
#
################################################################################
function Get-Converter
{
    Param
    (
        [parameter(Mandatory=$true)]
        [AttributeDefinition]$AttributeDefinition
    )

    $fileTimeAttributes = (
        "badPasswordTime",
        "lastLogonTimeStamp",
        "pwdLastSet",
        "accountExpires",
        "lastLogon",
        "lastPwdSet",
        "lockoutTime",
        "msDS-LastSuccessfulInteractiveLogonTime"
        )

    function Get-ConverterInstance
    {
        Param
        (
            [parameter(Mandatory=$true)]
            [string]$Converter,
            [parameter(Mandatory=$false)]
            $Param1
        )
        if((Get-Variable -Scope Script $Converter).Value -eq $null)
        {
            Write-Log -LogString ("Instantiating and returning IValueConverter: [" + $Converter + "][" + $AttributeDefinition.Attribute + "]") -Severity "Debug"
            (Get-Variable -Scope Script $Converter).Value = New-Object -TypeName $Converter -ArgumentList $Param1
            return (Get-Variable -Scope Script $Converter).Value
        }
        else
        {
            Write-Log -LogString ("Returning IValueConverter: [" + $Converter + "][" + $AttributeDefinition.Attribute + "]") -Severity "Debug"
            return (Get-Variable -Scope Script $Converter).Value
        }
    }

    if($AttributeDefinition.IgnoreConverter)
    {
        Write-Log -LogString ("Ignoring converter for attribute '" + $AttributeDefinition.Attribute + "'") -Severity "Debug"
        return $null
    }

    if($fileTimeAttributes -Contains ($AttributeDefinition.Attribute))
    {
        return Get-ConverterInstance -Converter "FileTimeConverter" -Param1 $script:dateTimeFormat
    }
    elseif($AttributeDefinition.Attribute -eq "managedBy")
    {
        return Get-ConverterInstance -Converter "managedByConverter"
    }
    elseif($AttributeDefinition.IsSingleValued -and $AttributeDefinition.Syntax -eq "GeneralizedTime")
    {
        return Get-ConverterInstance -Converter "DateFormatConverter" -Param1 $script:dateTimeFormat
    }
    elseif(!$AttributeDefinition.IsSingleValued)
    {
        return Get-ConverterInstance -Converter "ADPropertyValueCollectionConverter"
    }
    elseif($AttributeDefinition.Attribute -eq "msExchRemoteRecipientType")
    {
        return Get-ConverterInstance -Converter "msExchRemoteRecipientTypeConverter"
    }
    elseif($AttributeDefinition.Attribute -eq "msExchRecipientDisplayType")
    {
        return Get-ConverterInstance -Converter "msExchRecipientDisplayTypeConverter"
    }
}
###############################################################################
#
#.SYNOPSIS
#    Returns default attribute definitions for computer or user objects
#    These are only used on first run or if settings.xml is missing
#
#.PARAMETER Computer
#    Returns default attribute definitions for computer objects
#
#.PARAMETER User
#    Returns default attribute definitions for user objects
#
###############################################################################
function Get-DefaultAttributeDefinitions
{
    Param
    (
        [Parameter(Mandatory=$false)]
        [switch]$Computer,
        [Parameter(Mandatory=$false)]
        [switch]$User
    )

    function Get-DefaultComputerAttributeDefinitions
    {
        $computerAttributes = New-Object System.Collections.Generic.List[AttributeDefinition]
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Name",                "name",                   $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Company",             "company",                $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Division",            "division",               $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department",          "department",             $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department Number",   "departmentNumber",       $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Location",            "location",               $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Description",         "description",            $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last user",           "info",                   $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Computer model",      "employeeType",           $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Operating system",    "operatingSystem",        $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "OS version",          "operatingSystemVersion", $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last logged on",      "lastLogonTimestamp",     $false,  "DataGrid"))        
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Created",             "whenCreated",            $false,  "DataGrid"))
        return $computerAttributes
    }

    function Get-DefaultUserAttributeDefinitions
    {
        $userAttributes = New-Object System.Collections.Generic.List[AttributeDefinition]
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Name",                "name",                                    $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Company",             "company",                                 $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department",          "department",                              $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Title",               "title",                                   $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Description",         "description",                             $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Interactive logon",   "msDS-LastSuccessfulInteractiveLogonTime", $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last logged on",      "lastLogonTimestamp",                      $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "UPN",                 "userPrincipalName",                       $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Created",             "whenCreated",                             $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Password last set",   "pwdLastSet",                              $false,  "DetailsPane"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last bad password",   "badPasswordTime",                         $false,  "DetailsPane"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Bad password count",  "badPwdCount",                             $false,  "DetailsPane"))
        return $userAttributes
    }

    if($Computer)
    {
        return Get-DefaultComputerAttributeDefinitions
    }
    elseif($User)
    {
        return Get-DefaultUserAttributeDefinitions
    }

}
###############################################################################
#.SYNOPSIS
#    Returns a TreeViewItem with children containing the domains OU structure.
#    The ADOrganizationalUnit object is stored in the items tag
#    The root items tag will contain a ADDomain object
#
#.PARAMETER SearchString
#    Will return results based on SearchString
#
###############################################################################
function Get-DomainTree
{
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory=$false)]
        [string]$SearchString
    )

    function Get-OUTreeChildItems()
    {
        [CmdletBinding()]
        Param
        (
            [parameter(Mandatory=$true)]
            [System.Windows.Controls.TreeViewItem] $ParentItem
        )

        foreach($ou in Get-ADOrganizationalUnit -Filter * -SearchBase $ParentItem.Tag -SearchScope OneLevel)
        {
            $script:ouCount++
            $treeViewItem = New-Object System.Windows.Controls.TreeViewItem
            $treeViewItem.Tag = $ou
            $treeViewItem.Header = $ou.Name
            Get-OUTreeChildItems -ParentItem $treeViewItem
            $ParentItem.Items.Add($treeViewItem) | Out-Null
        }
    }

    Write-Log -LogString "Loading Organizational units..." -Severity "Informational"
    $startTime = (Get-Date)
    try
    {
        $rootItem = New-Object System.Windows.Controls.TreeViewItem
        $domain = Get-ADDomain
        $rootItem.Header = $domain.NetBIOSName
        $rootItem.Tag = $domain
        $script:ouCount = 0
        if([string]::IsNullOrWhiteSpace($SearchString))
        {
            Get-OUTreeChildItems -ParentItem $rootItem
        }
        else
        {
            $SearchString = ("*" + $SearchString + "*")
            Write-Log -LogString ("Loading Organizational units matching '" + $SearchString + "'") -Severity "Informational"
            foreach($ou in Get-ADOrganizationalUnit -Filter { Name -like $SearchString } )
            {
                $script:ouCount++
                Write-Log -LogString ("Found " + $ou.Name) -Severity "Debug"
                $treeViewItem = New-Object System.Windows.Controls.TreeViewItem
                $treeViewItem.Tag = $ou
                $treeViewItem.Header = $ou.Name
                $rootItem.Items.Add($treeViewItem) | Out-Null
            }
            $rootItem.Header = ($rootItem.Header + " (" + $script:ouCount.ToString() + ") search results")
        }
        Write-Log -LogString ($script:ouCount.ToString() + " units found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Critical"
    }
    $rootItem.IsExpanded = $true
    return $rootItem
}
###############################################################################
#.SYNOPSIS
#    Returns Groups from Active Directory
#
#.PARAMETER SearchString
#    Filters which groups to return based on SearchString
#
###############################################################################
function Get-Groups()
{
    Param
    (
        [parameter(Mandatory=$false)]
        [string]$SearchString
    )
    Write-Log -LogString "Loading Groups..." -Severity "Informational"
    $startTime = (Get-Date)
    try
    {
        if(!$SearchString)
        {
            $result = @(Get-ADGroup -Filter * | Sort-Object Name)
        }
        else
        {
            $filter = "*" + $SearchString + "*"
            $result = @(Get-ADGroup -Filter { Name -like $filter } | Sort-Object Name)
        }

        Write-Log -LogString ($result.Count.ToString() + " groups found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
        return ,$result
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Critical"
    }

}
###############################################################################
#
#.SYNOPSIS
#    Returns the LAPS password from Active Directory
#
#.PARAMETER Computer
#    Will return the LAPS password for this computer
#
###############################################################################
function Get-LapsPassword
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADComputer]$Computer        
    )
    Write-Log -LogString ("Trying to get LAPS password for computer [" + $Computer.Name + "]") -Severity "Notice"
    
    $lapspwd = Get-ADComputer -Identity $Computer.DistinguishedName -Properties ms-mcs-admpwd | Select-Object -Expand ms-mcs-admpwd        
    if(![string]::IsNullOrEmpty($lapspwd))
    {
        return $lapspwd
    }
    else
    {        
        try
        {
            $lapspwd = Get-ADComputer -Credential (Get-Credential -Message "Password looks empty, you probably don't have permissions to the confidential attribute. Specify different account with sufficient permissions to try again") -Identity $Computer.DistinguishedName -Properties ms-mcs-admpwd | Select-Object -Expand ms-mcs-admpwd            
            if(![string]::IsNullOrEmpty($lapspwd))
            {
                return $lapspwd
            }
            else
            {
                [System.Windows.MessageBox]::Show("The retrieved password is empty, you might have insufficient permissions to read the attribute: ", "Error",'Ok','Error') | Out-Null
            }
            
        }
        catch
        {
            Write-Log -LogString ("Exception retrieving LAPS password: " + $_.Exception.Message) -Severity "Error"
            [System.Windows.MessageBox]::Show("Exception retrieving LAPS password: " + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
            return $null
        }                     
    }    
}
###############################################################################
#.SYNOPSIS
#    Returns searchbase from either OU Browser or Groups Listbox (depending on which has a selected item)
#
###############################################################################
function Get-SearchBase
{
    if($script:MainWindow.tvOUBrowser.SelectedItem)
    {
        return $script:MainWindow.tvOUBrowser.SelectedItem.Tag
    }
    elseif($script:MainWindow.lstBoxGroups.SelectedIndex -gt -1)
    {
        return $script:MainWindow.lstBoxGroups.SelectedItem
    }
}
###############################################################################
#.SYNOPSIS
#    Returns searchscope from either OU Browser or Groups Listbox (depending on which has a selected item)
#
###############################################################################
function Get-SearchScope
{
    if((Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or (Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
    {
        if($script:MainWindow.chkRecursiveOUSearch.IsChecked)
        {
            return "Subtree"
        }
        else
        {
            return "OneLevel"
        }
    }
    elseif((Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
    {
        if($script:MainWindow.chkRecursiveGroupSearch.IsChecked)
        {
            return "Subtree"
        }
        else
        {
            return "OneLevel"
        }
    }
}
###############################################################################
#
#.SYNOPSIS
#    Returns the currently selected object (from the currently selected DataGrid)
#    Paremeters can be specified to override and only return selected object of
#    specific type
#
#.PARAMETER Computer
#    If specified, will return the currently selected computer object
#
#.PARAMETER User
#    If specified, will return the currently selected user object
#
###############################################################################
function Get-SelectedObject
{
    Param
    (
        [Parameter(Mandatory=$false)]
        [switch]$Computer,
        [Parameter(Mandatory=$false)]
        [switch]$User
    )
    if($Computer)
    {
        return $script:MainWindow.computersDataGrid.SelectedItem
    }
    elseif($User)
    {
        return $script:MainWindow.usersDataGrid.SelectedItem
    }
    else
    {
        if($script:MainWindow.tabItemComputers.IsSelected -eq $true)
        {
            $dg = $script:MainWindow.computersDataGrid
        }
        elseif($script:MainWindow.tabItemUsers.IsSelected -eq $true)
        {
            $dg = $script:MainWindow.usersDataGrid
        }
        return $dg.SelectedItem
    }
}
###############################################################################
#.SYNOPSIS
#    Returns a Settings object.
#
#.DESCRIPTION
#    Returns a Settings object.
#    It will first try to import from file .\settings.xml
#    If this fails, it will generate a new Settings object
#    with default values.
#
###############################################################################
function Get-Settings
{
    try
    {
        $settingsImport = Import-Clixml -Path $script:settingsFile
        Write-Log -LogString "Settings loaded from file. Generating Settings object..." -Severity "Notice"
        $settings = New-Object Settings
        foreach ($Property in ($settingsImport | Get-Member -MemberType Property))
        {
            if($Property.Name -eq "ComputerAttributeDefinitions" -or $Property.Name -eq "UserAttributeDefinitions")
            {
                foreach($importedDefinition in $settingsImport.$($Property.Name))
                {
                    $attributeDefinition = New-Object AttributeDefinition

                    foreach ($defProperty in ($importedDefinition | Get-Member -MemberType Property))
                    {
                        $attributeDefinition.$($defProperty.Name) = $importedDefinition.$($defProperty.Name)
                    }
                    $settings.$($Property.Name).Add($attributeDefinition)
                }
            }
            else
            {
                $settings.$($Property.Name) = $settingsImport.$($Property.Name)
            }
        }
        
        #        
        # Check sanity against schema as a precaution...
        #
        if(Resolve-AttributeDefinitionsFromSchema -ComputerAttributeDefinitions $settings.ComputerAttributeDefinitions -UserAttributeDefinitions $settings.UserAttributeDefinitions)
        {            
            Write-Log -LogString ("Returning Settings object from file") -Severity "Debug"
            return $settings
        }
        else
        {
            Write-Log -LogString "Unable to get default settings... script will terminate. Press enter to exit." -Severity "Critical"
            [System.Windows.MessageBox]::Show("Unable to get default settings... script will terminate." + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
            Read-Host
            exit
        }                
    }
    catch
    {
        #
        # Ok, something went wrong loading settings file or it does not exist. Generate new settings and as an extra
        # precaution, pass the definitions through Resolve-AttributeDefintionsFromSchema
        #
                
        if(($_.Exception.HResult -eq -2147024893) -or ($_.Exception.HResult -eq -2147024894))
        {
            #
            # Path not found, don't write red error to console
            #                
            Write-Log -LogString "Unable to load settings from file, using defaults" -Severity "Notice"    
        }
        else
        {
            #
            # Something other then a missing file, log an error!
            #
            Write-Log -LogString ("Reading settings file failed with error: " + $_.Exception.Message) -Severity "Error"
            Write-Log -LogString "Unable to load settings from file, using defaults" -Severity "Notice"
        }
                
        $newSettings = New-Object Settings
        $newSettings.ComputerAttributeDefinitions = Get-DefaultAttributeDefinitions -Computer
        $newSettings.UserAttributeDefinitions = Get-DefaultAttributeDefinitions -User

        if(Resolve-AttributeDefinitionsFromSchema -ComputerAttributeDefinitions $newSettings.ComputerAttributeDefinitions -UserAttributeDefinitions $newSettings.UserAttributeDefinitions)
        {
            Write-Log -LogString "Failed to get settings object from file, returning default settings object" -Severity "Notice"
            return $newSettings
        }
        else
        {
            Write-Log -LogString "Unable to get default settings... script will terminate. Press enter to exit." -Severity "Critical"
            [System.Windows.MessageBox]::Show("Unable to get default settings... script will terminate." + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
            Read-Host
            exit
        }
    }
}
###############################################################################
#.SYNOPSIS
#    Checks AD schema and sanitizes user provided AttributeDefinitions
#
#.DESCRIPTION
#    Checks AD schema and sanitizes user provided AttributeDefinitions
#    If user has set an attribute as editable with Syntax other than DirectoryString
#    We will override and set it as not editable.
#
#.PARAMETER ComputerAttributeDefinitions
#    ComputerAttributeDefinitions to check
#
#.PARAMETER UserAttributeDefinitions
#    UserAttributeDefinitions to check
#
#.NOTES
#    Some special handling here for 'description' attribute. It's marked as multivalued in schema, when it's mostly not...
#    From MSDN:
#    Remarks
#    The description attribute is implemented as a multi-valued attribute in the schema for the cases where that is allowed. For an object that is not a SAM
#    managed class, the description is multi-valued. For an attribute that is a SAM managed class, the description attribute is single-valued. SAM managed
#    classes are for things like security principals so, if you have, for example, a container, or a class of your own, the schema will let you use multiple values.
#    This behavior of the description attribute is for backward compatibility with earlier operating systems because the attribute existed in the SAM APIs
#    before AD existed.
#    https://msdn.microsoft.com/en-us/library/ms675492%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
#
###############################################################################
function Resolve-AttributeDefinitionsFromSchema
{
    Param
    (
        [parameter(Mandatory=$true)]
        [System.Collections.Generic.List[AttributeDefinition]]$ComputerAttributeDefinitions,
        [parameter(Mandatory=$true)]
        [System.Collections.Generic.List[AttributeDefinition]]$UserAttributeDefinitions
    )

    function Resolve()
    {
        Param
        (
            [parameter(Mandatory=$true)]
            [System.Collections.Generic.List[AttributeDefinition]]$AttributeDefinitions,
            [parameter(Mandatory=$true)]
            [string]$ObjectCategory
        )

        foreach($attrDef in $AttributeDefinitions)
        {
            try
            {
                $schemaProperty = $schema.FindProperty($attrDef.Attribute)
                #Check if Single or MultiValue
                if($attrDef.Attribute -eq "description")
                {
                    Write-Log -LogString ("[" + $ObjectCategory + "] '" + $attrDef.Attribute + "' setting as SingleValued <--- *** DEFINITION CHANGED ***") -Severity "Debug"
                    $attrDef.IsSingleValued = $true
                }
                else
                {
                    $attrDef.IsSingleValued = $schemaProperty.IsSingleValued

                }
                #Get syntax
                $attrDef.Syntax = $schemaProperty.Syntax

                #Finally set AttributeDefinition to read-only if Syntax is not DirectoryString
                if($attrDef.Syntax -ne "DirectoryString" -and $attrDef.IsEditable -eq $true)
                {
                    $attrDef.IsEditable = $false
                    Write-Log -LogString ("[" + $ObjectCategory + "] '" + $attrDef.Attribute +  "' Syntax:" + $attrDef.Syntax + " setting as Read-Only <--- *** DEFINITION CHANGED ***") -Severity "Debug"
                }
            }
            catch
            {
                Write-Log -LogString $_.Exception.Message -Severity "Critical"
                [System.Windows.MessageBox]::Show("Cannot resolve attribute '" + $attrDef.Attribute + "' from schema. " + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
                return $false
            }

            Write-Log -LogString ("[" + $ObjectCategory + "] '" + $attrDef.Attribute +  "' Syntax:" + $attrDef.Syntax + " IsSingleValued:" + $attrDef.IsSingleValued) -Severity "Debug"
        }
        return $true
    }

    $schema = [DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
    $startTime = Get-Date
    Write-Log -LogString "Resolving AttributeDefinitions from schema..." -Severity "Notice"
    $computerResolveResult = Resolve -AttributeDefinition $ComputerAttributeDefinitions -ObjectCategory "computer"
    $userResolveResult = Resolve -AttributeDefinition $UserAttributeDefinitions -ObjectCategory "user"

    Write-Log -LogString ("Computer attributes resolve result:: " + $computerResolveResult.ToString()) -Severity "Debug"
    Write-Log -LogString ("User attributes resolve result: " + $userResolveResult.ToString()) -Severity "Debug"

    Write-Log -LogString ("Resolved in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Notice"

    if($computerResolveResult -eq $true -and $userResolveResult -eq $true)
    {
        return $true
    }
    else
    {
        return $false
    }

}
###############################################################################
#.SYNOPSIS
#    Commits attribute changes to AD
#
#.DESCRIPTION
#    It will iterate over all writeable attributes of the selected object type.
#    It will compare existing values with values in the Comboboxes. It can
#    connect the ComboBoxes to specific attributes since the ComboBox Tag contains
#    the string representation of the attribute. If values don't match they will
#    be updated in the object instance. And $commitChanges will be set to $true.
#
#    When iteration is finished, if $commitChanges is $true, changes will be
#    commited by either Set-ADComputer or Set-ADUser with -Instance $TargetObject
#
#.PARAMETER TargetObject
#    The instance of the object to set values on.
#
#.NOTES
#   General notes
#
###############################################################################
function Set-AttributeValues
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [System.Object]$TargetObject
    )

    $writeableAttributes = ""
    if($TargetObject.GetType() -eq [Microsoft.ActiveDirectory.Management.ADComputer])
    {
        $writeableAttributes = @($script:settings.ComputerAttributeDefinitions | Where-Object {$_.IsEditable -eq $true} )
        $typePrefix = "Computer"
        $stkEditableAttributes = $script:MainWindow.stkEditableComputerAttributes
    }
    elseif($TargetObject.GetType() -eq [Microsoft.ActiveDirectory.Management.ADUser])
    {
        $writeableAttributes = @($script:settings.UserAttributeDefinitions | Where-Object {$_.IsEditable -eq $true} )
        $typePrefix = "User"
        $stkEditableAttributes = $script:MainWindow.stkEditableUserAttributes
    }
    else
    {
        Write-Log -LogString "Cannot commit, unknown target type" -Severity "Critical"
        [System.Windows.MessageBox]::Show("Cannot commit, unknown target type", "Error",'Ok','Error') | Out-Null
        return
    }

    $commitChanges = $false
    $commitCount = 0
    foreach($attr in $writeableAttributes)
    {
        $newValue = ($stkEditableAttributes.Children | Where-Object {$_.GetType() -eq [System.Windows.Controls.ComboBox] -and $_.Tag -eq $attr.Attribute}).Text
        if([string]::IsNullOrWhiteSpace($newValue))
        {
            $newValue = $null
        }

        if($attr.IsSingleValued)
        {
            $existingValue = $TargetObject.($attr.Attribute)
        }
        else
        {
            $existingValue = @($TargetObject.($attr.Attribute))
            if($newValue)
            {
                $newValue = @($newValue.Split(';', [System.StringSplitOptions]::RemoveEmptyEntries))
            }
        }

        $verboseStr = ("[" + $TargetObject.Name + "] Attribute [" + $attr.Attribute + "]").PadRight(54)
        if((!$existingValue -and $newValue) -or ($existingValue -and !$newValue) -or ($existingValue -and $newValue -and (Compare-Object -CaseSensitive $newValue $existingValue)))
        {
            if($newValue -eq $null)
            {
                Write-Log -LogString ($verboseStr + " pending action [clear]") -Severity "Notice"
                $commitCount++
            }
            else
            {
                Write-Log -LogString ($verboseStr + " pending action [new value]") -Severity "Notice"
                $commitCount++
            }
            $TargetObject.($attr.Attribute) = $newValue
            $commitChanges = $true
        }
        else
        {
            Write-Log -LogString ($verboseStr + " pending action [no changes]") -Severity "Notice"
        }
    }

    if($commitChanges)
    {
        try
        {
            if($TargetObject.GetType() -eq [Microsoft.ActiveDirectory.Management.ADComputer])
            {
                Set-ADComputer -Instance $TargetObject
                $script:MainWindow.computersDataGrid.Items.Refresh()
            }
            elseif($TargetObject.GetType() -eq [Microsoft.ActiveDirectory.Management.ADUser])
            {
                Set-ADUser -Instance $TargetObject
                $script:MainWindow.usersDataGrid.Items.Refresh()
            }
            Write-Log -LogString ("[" + $TargetObject.Name + "] " + $commitCount.ToString() + " change(s) commited...") -Severity "Notice"
        }
        catch
        {
            Write-Log -LogString $_.Exception.Message -Severity "Critical"
            [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error') | Out-Null
        }
    }
    else
    {
        Write-Log -LogString ("[" + $TargetObject.Name + "] " + $commitCount.ToString() + " change(s) commited...") -Severity "Notice"
    }    
}
function Show-DebugWindow
{    
    Param
    (
        [parameter(Mandatory=$True)]
        $ItemsSource
    )

    #
    # Setup Window
    #
    $debugWindow = @{}
    $debugWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamldebugWindow))
    $debugWindow.Window.Title = "Debug view"
    $style = ($debugWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamldebugWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        #Write-Log -LogString ("Adding " + $guiObject.Name + " to $debugWindow") -Severity "Debug"
        $debugWindow.$($guiObject.Name) = $debugWindow.Window.FindName($guiObject.Name)
    }

    $debugWindow.dg.ItemsSource = $ItemsSource
   
    $debugWindow.Window.ShowDialog()

}
###############################################################################
#.SYNOPSIS
#    Exports objects from currently active DataGrid to a csv file.
#
#.DESCRIPTION
#    Exports objects from currently active DataGrid to a csv file.
#    Objects are iterated over and values are passed to a PSObject to get the header
#    names from attribute FriendlyName. And to get converted value for certain attributes.
#
###############################################################################
function Show-ExportWindow
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $AttributeDefinitions,
        [Parameter(Mandatory=$true)]
        $Source
    )

    #
    # Setup Window
    #
    $exportWindow = @{}
    $exportWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamlExportWindow))
    $exportWindow.Window.Title = "Export"
    $style = ($exportWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamlExportWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        Write-Log -LogString ("Adding " + $guiObject.Name + " to $exportWindow") -Severity "Debug"
        $exportWindow.$($guiObject.Name) = $exportWindow.Window.FindName($guiObject.Name)
    }

    $exportWindow.Window.DataContext = $AttributeDefinitions

    function Export-Inventory
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if($exportWindow.rdBtnCSV.IsChecked)
        {
            $SaveFileDialog.Filter = “Text files (*.csv)|*.csv|All files (*.*)|*.*”
        }
        else
        {
            $SaveFileDialog.Filter = “XLSX files (*.xlsx)|*.xlsx|All files (*.*)|*.*”
        }

        $SaveFileDialog.ShowDialog()
        $SaveFileDialog.Filename

        if($SaveFileDialog.Filename -ne "")
        {
            $exportObjects = New-Object System.Collections.ObjectModel.ObservableCollection[object]
            foreach($adObj in  $Source)
            {
                $exportObj = New-Object -TypeName PSObject
                foreach($attr in $exportWindow.lstBoxColumnsToInclude.SelectedItems)
                {
                    if($adObj.($attr.Attribute))
                    {
                        if($converter = Get-Converter -AttributeDefinition $attr)
                        {
                            $exportObj | Add-Member -type NoteProperty -Name $attr.FriendlyName -value $converter.Convert($adObj.($attr.Attribute), $null, $null, $null)
                        }
                        else
                        {
                            $exportObj | Add-Member -type NoteProperty -Name $attr.FriendlyName -value ($adObj.($attr.Attribute))
                        }
                    }
                    else # Object has no value in this attribute, so add attribute but with null value (mainly to avoid null errors while sorting or grouping...)
                    {
                        $exportObj | Add-Member -type NoteProperty -Name $attr.FriendlyName -value $null
                    }
                }
                $exportObjects.Add($exportObj) | Out-Null
            }
            $delimiter = $exportWindow.txtBoxDelimiter.Text[0]
            if([string]::IsNullOrWhiteSpace($delimiter))
            {
                Write-Log -LogString "User defined delimiter is null or empty. Setting delimiter to ';'" -Severity "Debug"
                $delimiter = ";"
            }
            else
            {
                Write-Log -LogString ("Delimiter is '" + $delimiter + "'") -Severity "Debug"
            }

            # Hack, otherwise columns will be sorted by order in which they where selected in ListBox
            # This will sort them in order of Attribute Definitions
            [System.Collections.Generic.List[AttributeDefinition]] $AttributesToExport = New-Object System.Collections.Generic.List[AttributeDefinition]
            foreach($def in $exportWindow.lstBoxColumnsToInclude.Items)
            {
                if($exportWindow.lstBoxColumnsToInclude.SelectedItems.Contains($def))
                {
                   $AttributesToExport.Add($def)
                }
            }

            #
            # Do the export
            #
            if($exportWindow.rdBtnCSV.IsChecked) # Do CSV export
            {
                try
                {
                    Write-Log -LogString "Exporting to CSV" -Severity "Informational"
                    $exportObjects = $exportObjects | Sort-Object -Property $exportWindow.lstBoxSortBy.SelectedItem.FriendlyName
                    $exportObjects | Select-Object ($AttributesToExport).FriendlyName | Export-CSV -Path $SaveFileDialog.filename -Delimiter $delimiter -NoTypeInformation -Encoding UTF8
                    Invoke-Item $SaveFileDialog.filename
                }
                catch
                {                    
                    [System.Windows.MessageBox]::Show("Export in CSV format failed. Se console for detailed error message.","Error",'Ok','Error') | Out-Null
                    Write-Log -LogString ("CSV export failed with error: " + $_.Exception.Message) -Severity "Error"                    
                }                  
            }
            else # Do EXCEL export
            {
                Write-Log -LogString "Exporting to XLSX" -Severity "Informational"

                try
                {
                    Write-Log -LogString "Loading Excel ComObject..." -Severity "Debug"
                    $excel = New-Object -ComObject Excel.Application

                    $workbook = $excel.Workbooks.Add()
                    $sheet = $workbook.Sheets[1]
                    $rowIndex = 1

                    #
                    # Add summary
                    #
                    Write-Log -LogString "Creating summary..." -Severity "Informational"
                    $sheet.Cells.Item($rowIndex,1) = "Created: "
                    $sheet.Cells.Item($rowIndex,2) = ((Get-Date).ToString() + " by " + $env:USERNAME + " on " + $env:COMPUTERNAME)
                    $sheet.Cells.Item($rowIndex+1,1) = ("Objects: " + $exportObjects.Count)

                    if($Source[0].GetType() -eq [Microsoft.ActiveDirectory.Management.ADComputer])
                    {
                        $total = @($Source).Count.ToString()
                        $active = @($Source | Where-Object {$_.lastLogonTimeStamp -gt (Get-Date).AddDays(-$script:settings.ComputerInactiveLimit).ToFileTimeUtc() }).Count.ToString()
                        $passive = @($Source | Where-Object {$_.lastLogonTimeStamp -lt (Get-Date).AddDays(-$script:settings.ComputerInactiveLimit).ToFileTimeUtc() }).Count.ToString()
                    }
                    elseif($Source[0].GetType() -eq [Microsoft.ActiveDirectory.Management.ADUser])
                    {
                        $total = @($Source).Count.ToString()
                        $active = @($Source | Where-Object {$_.lastLogonTimeStamp -gt (Get-Date).AddDays(-$script:settings.UserInactiveLimit).ToFileTimeUtc() }).Count.ToString()
                        $passive = @($Source | Where-Object {$_.lastLogonTimeStamp -lt (Get-Date).AddDays(-$script:settings.UserInactiveLimit).ToFileTimeUtc() }).Count.ToString()
                    }

                    $sheet.Cells.Item($rowIndex+1,1) = ("Total: ")
                    $sheet.Cells.Item($rowIndex+1,2) = ($total)
                    $sheet.Cells.Item($rowIndex+1,2).HorizontalAlignment = -4131

                    $sheet.Cells.Item($rowIndex+2,1) = ("Active: ")
                    $sheet.Cells.Item($rowIndex+2,2) = ($active + " [" + [math]::Round($active / $total * 100) + "%]")
                    $sheet.Cells.Item($rowIndex+2,2).HorizontalAlignment = -4131
                    $sheet.Cells.Item($rowIndex+2,2).Font.ColorIndex = 10

                    $sheet.Cells.Item($rowIndex+3,1) = ("Passive: ")
                    $sheet.Cells.Item($rowIndex+3,2) = ($passive + " [" + [math]::Round($passive / $total * 100) + "%]")
                    $sheet.Cells.Item($rowIndex+3,2).HorizontalAlignment = -4131
                    $sheet.Cells.Item($rowIndex+3,2).Font.ColorIndex = 3

                    for($i = 1; $i -lt 6; $i++)
                    {
                        $mergeCells = $sheet.Range($sheet.Cells.Item($rowIndex,2), $sheet.Cells.Item($rowIndex,$AttributesToExport.Count + 1))
                        $mergeCells.Select()
                        $mergeCells.MergeCells = $true
                        $rowIndex++
                    }

                    function Add-Header
                    {
                        $index = 1
                        foreach($def in $AttributesToExport)
                        {
                            $sheet.Cells.Item($rowIndex,$index) = $def.FriendlyName
                            # Set cell color
                            $sheet.Cells.Item($rowIndex,$index).Interior.ColorIndex = 15
                            $index++
                        }
                    }

                    function Add-GroupHeader
                    {
                        Param
                        (
                            [Parameter(Mandatory=$true)]
                            [string]$GroupDescription
                        )
                        $sheet.Cells.Item($rowIndex,1) = $GroupDescription
                        # Merge cells
                        $mergeCells = $sheet.Range($sheet.Cells.Item($rowIndex,1), $sheet.Cells.Item($rowIndex,$AttributesToExport.Count))
                        $mergeCells.Select()
                        $mergeCells.MergeCells = $true
                        # Set cell color
                        $sheet.Cells.Item($rowIndex,1).Interior.ColorIndex = 16
                        $sheet.Cells.Item($rowIndex,1).HorizontalAlignment = -4108
                    }

                    #
                    # Sort objects
                    #
                    $sortedObjects = $exportObjects | Sort-Object -Property $exportWindow.lstBoxSortBy.SelectedItem.FriendlyName

                    #
                    # Group objects
                    #
                    if($exportWindow.lstBoxGroupBy.SelectedIndex -gt -1)
                    {
                        # Clear this list, all items should be in $sortedObjects
                        $exportObjects = New-Object System.Collections.ObjectModel.ObservableCollection[object]

                        $groupBy = $exportWindow.lstBoxGroupBy.SelectedItem.FriendlyName
                        Write-Log -LogString ("Generating grouped report, grouping by: '" + $groupBy + "'") -Severity "Informational"

                        #
                        # Need to filter out objects missing the $attr.Attribute property. Set-StrictMode will complain otherwise...
                        #
                        $groups = @($sortedObjects | Where-Object {Get-Member -InputObject $_ -Name $groupBy -Membertype Properties})

                        #$groups = @($groups.($groupBy) | Where-Object {$_ -ne $null} | Sort-Object | Get-Unique )  ## Get-Unique is case sensitive....
                        $groups = @($groups.($groupBy) | Where-Object {$_ -ne $null} | Sort-Object -Unique )

                        $count = 0
                        foreach($a in $groups)
                        {
                            Write-Log -LogString ("Exporting group: " + $a) -Severity "Debug"
                            # Add group delimiter
                            Add-GroupHeader -GroupDescription ("Group: '" + $a + "' " + @($sortedObjects | Where-Object {$_.$groupBy -ne $null} | Where-Object {$_.$groupBy -eq $a}).Count + " objects")
                            $rowIndex++
                            Add-Header
                            $rowIndex++
                            foreach($obj in @($sortedObjects | Where-Object {$_.$groupBy -ne $null} | Where-Object {$_.$groupBy -eq $a}))
                            {
                                $matrix = @()
                                foreach($attr in $AttributesToExport)
                                {
                                    $matrix = $matrix += $obj.($attr.FriendlyName) # This is inefficient since it will create a new array on every add...
                                }
                                $cellRange = $sheet.Range($sheet.Cells.Item($rowIndex,1), $sheet.Cells.Item($rowIndex,$AttributesToExport.Count))
                                $cellRange.value = $matrix
                                $rowIndex++

                                $count++
                                Write-Progress -activity ("Exporting group [" + $a + "]...") -status "Finished: $count of $($sortedObjects.Count)" -percentComplete (($count / $sortedObjects.Count)  * 100) -Id 1
                            }
                            Write-Log -LogString ("'" + $a + "' exported. Contains " + @($sortedObjects | Where-Object {$_.$groupBy -ne $null} | Where-Object {$_.$groupBy -eq $a}).Count + " objects") -Severity "Debug"

                            $rowIndex++ # Add empty row at end of group
                        }

                        #
                        # Export objects will nullvalue in groupBy attribute
                        #
                        Add-GroupHeader -GroupDescription ("Group: 'null' " + @($sortedObjects | Where-Object {$_.$groupBy -eq $null}).Count + " objects")
                        $rowIndex++
                        Add-Header
                        $rowIndex++
                        foreach($obj in @($sortedObjects | Where-Object {$_.$groupBy -eq $null}))
                        {
                            $matrix = @()
                            foreach($attr in $AttributesToExport)
                            {
                                $matrix = $matrix += $obj.($attr.FriendlyName) # This is inefficient since it will create a new array on every add...
                            }
                            $cellRange = $sheet.Range($sheet.Cells.Item($rowIndex,1), $sheet.Cells.Item($rowIndex,$AttributesToExport.Count))
                            $cellRange.value = $matrix
                            $rowIndex++

                            $count++
                            Write-Progress -activity "Exporting group [null]..." -status "Finished: $count of $($sortedObjects.Count)" -percentComplete (($count / $sortedObjects.Count)  * 100) -Id 1
                        }
                        Write-Log -LogString ("'null' exported. Contains " + @($sortedObjects | Where-Object {$_.$groupBy -eq $null}).Count + " objects") -Severity "Debug"
                    }
                    else
                    {
                        Write-Log -LogString "Generating non grouped export..." -Severity "Informational"
                        Add-Header
                        $rowIndex++
                        $count = 0
                        foreach($obj in @($sortedObjects))
                        {
                            $matrix = @()
                            foreach($attr in $AttributesToExport)
                            {
                                $matrix = $matrix += $obj.($attr.FriendlyName) # This is inefficient since it will create a new array on every add...
                            }
                            $cellRange = $sheet.Range($sheet.Cells.Item($rowIndex,1), $sheet.Cells.Item($rowIndex,$AttributesToExport.Count))
                            $cellRange.value = $matrix
                            $rowIndex++
                            $count++
                            Write-Progress -activity ("Exporting object [" + $obj.Name + "]...") -status "Finished: $count of $($sortedObjects.Count)" -percentComplete (($count / $sortedObjects.Count)  * 100) -Id 1
                        }
                    }

                    Write-Progress -Activity "Exporting objects..." -Completed -Id 1

                    $excel.Columns.AutoFit()
                    $excel.Visible = $true
                    $sheet.SaveAs($SaveFileDialog.filename)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
                }
                catch
                {                    
                    [System.Windows.MessageBox]::Show("Export in Excel format failed. Se console for detailed error message.","Error",'Ok','Error') | Out-Null
                    Write-Log -LogString ("Excel export failed with error: " + $_.Exception.Message) -Severity "Error"                    
                }                                                                
            }
        }
    }

    #
    # Select all attributes clicked
    #
    $exportWindow.btnSelectAll.add_Click({
        $exportWindow.lstBoxColumnsToInclude.SelectAll()
    })

    #
    # Ok button clicked
    #
    $exportWindow.btnOk.add_Click({
       Export-Inventory
    })

    #
    # Cancel button clicked
    #
    $exportWindow.btnCancel.add_Click({
        $exportWindow.Window.Close()
     })

    $exportWindow.Window.ShowDialog() | Out-Null
}
###############################################################################
#
#.SYNOPSIS
#    Shows a computer password in an independent Window
#
#.PARAMETER Password
#    Password to show
#
#.PARAMETER Hostname
#    Name of computer the password belongs to
#
###############################################################################
function Show-LapsPassword
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$Password,
        [Parameter(Mandatory=$true)]
        [string]$Hostname
    )

    #
    # Setup Window
    #
    $lapsWindow = @{}
    $lapsWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamllapsWindow))
    $lapsWindow.Window.Title = "LAPS Password"
    $style = ($lapsWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamllapsWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        Write-Log -LogString ("Adding " + $guiObject.Name + " to $lapsWindow") -Severity "Debug"
        $lapsWindow.$($guiObject.Name) = $lapsWindow.Window.FindName($guiObject.Name)
    }

    $lapsWindow.btnClose.add_Click({
        $lapsWindow.Window.Close()
    })

    $lapsWindow.btnShowPassword.add_Click({
        if($lapsWindow.txtPasswd.Text -ne $Password)
        {
            $lapsWindow.txtPasswd.Text = $Password
        }
        else
        {
            $lapsWindow.txtPasswd.Text = "********"       
        }
        
    })
    
    $lapsWindow.txtHostname.Text = $Hostname
    $lapsWindow.txtPasswd.Text = "********"     
    $lapsWindow.Window.ShowDialog()

}
###############################################################################
#.SYNOPSIS
#    Sets up GUI events and shows Window.
#.DESCRIPTION
#
#.NOTES
#    General notes
#
###############################################################################
function Show-MainWindow
{
    $script:MainWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamlMainWindow))
    $script:MainWindow.Window.Title = $script:appVersion
    $style = ($script:MainWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $script:xamlMainWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        $script:MainWindow.$($guiObject.Name) = $script:MainWindow.Window.FindName($guiObject.Name)
    }    

    #
    # OU Browser TreeViewItem has been selected (or unselected)
    #
    $script:MainWindow.tvOUBrowser.add_SelectedItemChanged({
        if($script:MainWindow.tvOUBrowser.SelectedItem)
        {
            # When an OU has been selected, unselect item in GroupListBox
            $script:MainWindow.lstBoxGroups.SelectedIndex = -1
            Update-DatagridsItemsSources
        }
    })

    #
    # Group ListBoxItem has been selected (or unselected)
    #
    $script:MainWindow.lstBoxGroups.add_SelectionChanged({
        if($script:MainWindow.lstBoxGroups.SelectedIndex -gt -1)
        {
            # When a group has been selected, unselect item in OU TreeView
            if($script:MainWindow.tvOUBrowser.SelectedItem)
            {
                $script:MainWindow.tvOUBrowser.SelectedItem.IsSelected = $false
            }
            Update-DatagridsItemsSources
        }
    })

    #
    # Filter groups button clicked
    #
    $script:MainWindow.btnFilterGroups.add_Click({
        $script:MainWindow.lstBoxGroups.ItemsSource = Get-Groups -SearchString $script:MainWindow.txtGroupFilter.Text
    })

    #
    # Reload groups button clicked
    #
    $script:MainWindow.btnReloadGroups.add_Click({
        $script:MainWindow.txtGroupFilter.Text = ""
        $script:MainWindow.lstBoxGroups.ItemsSource = Get-Groups
    })

    #
    # Key has been pressed in Group filter textbox
    #
    $script:MainWindow.txtGroupFilter.add_KeyDown({
        if ($args[1].Key -eq "Return")
        {
            $script:MainWindow.lstBoxGroups.ItemsSource = Get-Groups -SearchString $script:MainWindow.txtGroupFilter.Text
        }
        elseif($args[1].Key -eq "Escape")
        {
            $script:MainWindow.txtGroupFilter.Text = ""
        }
    })

    #
    # Filter OUs button clicked
    #
    $script:MainWindow.btnFilterOUs.add_Click({
        $script:MainWindow.tvOUBrowser.Items.Clear()
        $script:MainWindow.tvOUBrowser.Items.Add((Get-DomainTree -SearchString $script:MainWindow.txtOUFilter.Text)) | Out-Null
    })

    #
    # Reload OUs button clicked
    #
    $script:MainWindow.btnReloadOUs.add_Click({
        $script:MainWindow.txtOUFilter.Text = ""
        $script:MainWindow.tvOUBrowser.Items.Clear()
        $script:MainWindow.tvOUBrowser.Items.Add((Get-DomainTree)) | Out-Null
    })

    #
    # Key has been pressed in OU filter textbox.
    #
    $script:MainWindow.txtOUFilter.add_KeyDown({
        if ($args[1].Key -eq "Return")
        {
            $script:MainWindow.tvOUBrowser.Items.Clear()
            $script:MainWindow.tvOUBrowser.Items.Add((Get-DomainTree -SearchString $script:MainWindow.txtOUFilter.Text)) | Out-Null
        }
        elseif($args[1].Key -eq "Escape")
        {
            $script:MainWindow.txtOUFilter.Text = ""
        }
    })

    #
    # Computer/user TabItem selection changed
    #
    $script:MainWindow.tabControlDataGrids.add_SelectionChanged({
        # Set datacontext of ButtonBar to the currently selected DataGrid
        # SelectionChange event bubbles up from child items, so we need to check that event source is TabControl.
        if($args[1].OriginalSource.GetType() -eq [System.Windows.Controls.TabControl])
        {
            if($script:MainWindow.tabItemComputers.IsSelected)
            {
                Write-Log -LogString "Setting 'computersDataGrid' as DataContext for Button Bar" -Severity "Debug"
                $script:MainWindow.grdButtonBar.DataContext = $script:MainWindow.computersDataGrid
            }
            else
            {
                Write-Log -LogString "Setting 'usersDataGrid' as DataContext for Button Bar" -Severity "Debug"
                $script:MainWindow.grdButtonBar.DataContext = $script:MainWindow.usersDataGrid
            }
        }
    })

    
    #
    # DebugView button clicked. Show Debug Window
    #
    $script:MainWindow.btnDebugView.add_Click({
        if($script:MainWindow.tabItemComputers.IsSelected)
            {                
                if($script:MainWindow.computersDataGrid.ItemsSource -ne $null)
                {
                    Write-Log -LogString "Showing debug window with ItemsSource from Computers DataGrid" -Severity "Debug"
                    Show-DebugWindow -ItemsSource $script:MainWindow.computersDataGrid.ItemsSource                
                }
                else
                {
                    Write-Log -LogString "Computer ItemsSource it empty, no debug data to show!" -Severity "Warning"
                    [System.Windows.MessageBox]::Show("Computer ItemsSource it empty, no debug data to show!", "Error",'Ok','Error') | Out-Null
                }
                
            }
            else
            {
                if($script:MainWindow.usersDataGrid.ItemsSource -ne $null)
                {
                    Write-Log -LogString "Showing debug window with ItemsSource from Users DataGrid" -Severity "Debug"
                    Show-DebugWindow -ItemsSource $script:MainWindow.usersDataGrid.ItemsSource                
                }
                else
                {
                    Write-Log -LogString "User ItemsSource it empty, no debug data to show!" -Severity "Warning"
                    [System.Windows.MessageBox]::Show("User ItemsSource it empty, no debug data to show!", "Error",'Ok','Error') | Out-Null
                }
            }        
    })
    
    #
    # Settingsbutton clicked. Show Settings Window.
    #
    $script:MainWindow.btnSettings.add_Click({
        Show-SettingsWindow
    })

    #
    # Filter button clicked
    #
    $script:MainWindow.btnFilter.add_Click({
        Update-DatagridsItemsSources  -SearchString $script:MainWindow.txtFilter.Text
    })

    #
    # Enter/Return pressed in filter TextBox
    #
    $script:MainWindow.txtFilter.add_KeyDown({
        if ($args[1].Key -eq "Return")
        {
            Update-DatagridsItemsSources -SearchString $script:MainWindow.txtFilter.Text
        }
        elseif($args[1].Key -eq "Escape")
        {
            $script:MainWindow.txtFilter.Text = ""
        }
    })

    #
    # Show/Hide console button clicked
    #
    $script:MainWindow.btnShowHideConsole.add_Click({
        $consolePtr = [Console.Window]::GetConsoleWindow()
        if($script:MainWindow.btnShowHideConsole.IsChecked)
        {
            [Console.Window]::ShowWindow($consolePtr, 1)
            Write-Log -LogString "Showing console... *** Warning! *** Closing console window will terminate the script. Use togglebutton to hide it again." -Severity "Warning"
        }
        else
        {
            Write-Log -LogString "Hiding console..." -Severity "Informational"
            [Console.Window]::ShowWindow($consolePtr, 0)
        }
    })

    $script:MainWindow.Window.add_Loaded({
        $consolePtr = [Console.Window]::GetConsoleWindow()
        [Console.Window]::ShowWindow($consolePtr, 0)
    })

    #
    # Commit computer changes button clicked
    #
    $script:MainWindow.btnCommitComputerChanges.add_Click({
        Set-AttributeValues -TargetObject (Get-SelectedObject -Computer)
    })

    #
    # Commit user changes button clicked
    #
    $script:MainWindow.btnCommitUserChanges.add_Click({
        Set-AttributeValues -TargetObject (Get-SelectedObject -User)
    })

    #
    # Export button clicked
    #
    $script:MainWindow.btnExportData.add_Click({        
        if($script:MainWindow.tabItemComputers.IsSelected)
        {
            Show-ExportWindow -Source $script:MainWindow.computersDataGrid.ItemsSource -AttributeDefinition $script:settings.ComputerAttributeDefinitions
        }
        else
        {
            Show-ExportWindow -Source $script:MainWindow.usersDataGrid.ItemsSource -AttributeDefinition $script:settings.UserAttributeDefinitions
        }

    })

    #
    # Get LAPS password button clicked
    #
    $script:MainWindow.btnGetLapsPassword.add_Click({        
        $hostname = $script:MainWindow.computersDataGrid.SelectedItem.Name
        $pwd = Get-LapsPassword -Computer $script:MainWindow.computersDataGrid.SelectedItem
        if($pwd -ne $null)
        {
            Show-LapsPassword -Password $pwd -Hostname $hostname
        }
        else
        {
            Write-Log -LogString "Failed to retrieve LAPS password" -Severity "Notice"
        }                
    })

    #
    # RDP ContextMenuItem clicked
    #
    $script:MainWindow.ctxRDP.add_Click({
        if(Get-SelectedObject)
        {
            Write-Log -LogString ("Connecting with RDP to [" + (Get-SelectedObject).Name + "]") -Severity "Notice"
            &mstsc.exe /V: (Get-SelectedObject).Name
        }

    })    

    #
    # MSRA ContextMenuItem clicked
    #
    $script:MainWindow.ctxMSRA.add_Click({
        if(Get-SelectedObject)
        {
            Write-Log -LogString ("Offering remote assistance to [" + (Get-SelectedObject).Name + "]") -Severity "Notice"            
            &msra.exe /offerra (Get-SelectedObject).Name
        }
    })

    if($script:settings.OnStartLoadGroups)
    {
        $script:MainWindow.lstBoxGroups.ItemsSource = Get-Groups
    }
    if($script:settings.OnStartLoadOrganizationalUnits)
    {
        $script:MainWindow.tvOUBrowser.Items.Add((Get-DomainTree)) | Out-Null
    }

    $script:MainWindow.Window.ShowDialog() | Out-Null
}


###############################################################################
#.SYNOPSIS
#    Shows the Settings window.
#
#.NOTES
#
###############################################################################
function Show-SettingsWindow
{
    $settingsWindow = @{}
    $settingsWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamlSettingsWindow))
    $settingsWindow.Window.Title = "Settings"
    $style = ($settingsWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamlSettingsWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        $settingsWindow.$($guiObject.Name) = $settingsWindow.Window.FindName($guiObject.Name)
    }

    # Moves definition up in list
    function Move-Up
    {
        Param(
        [Parameter(Mandatory=$true)]
        $SourceList,
        [Parameter(Mandatory=$true)]
        $Item
        )

        $index = $SourceList.IndexOf($Item)
        if($index -gt 0)
        {
            $SourceList.Remove($Item)
            $SourceList.Insert(($Index - 1), $Item)
        }

    }

    # Moves definition down in list
    function Move-Down
    {
        Param(
        [Parameter(Mandatory=$true)]
        $SourceList,
        [Parameter(Mandatory=$true)]
        $Item
        )
        $index = $SourceList.IndexOf($Item)
        if($index -lt ($SourceList.Count - 1))
        {
            $SourceList.Remove($Item)
            $SourceList.Insert(($Index + 1), $Item)
        }

    }

    #
    # Add Computer definition clicked
    #
    $settingsWindow.btnAddComputerAttributeDefinition.add_Click({
        $settingsUnconfirmed.ComputerAttributeDefinitions.Add((New-Object AttributeDefinition))
        $settingsWindow.dgComputerAttributes.Items.Refresh()
    })

    #
    # Remove Computer definition clicked
    #
    $settingsWindow.btnRemoveComputerAttributeDefinition.add_Click({
        $settingsUnconfirmed.ComputerAttributeDefinitions.Remove($settingsWindow.dgComputerAttributes.SelectedItem)
        $settingsWindow.dgComputerAttributes.Items.Refresh()
    })

    #
    # Move Computer definition up clicked
    #
    $settingsWindow.btnUpComputerAttributeDefinition.add_Click({
        Move-Up -SourceList $settingsUnconfirmed.ComputerAttributeDefinitions -Item $settingsWindow.dgComputerAttributes.SelectedItem
        $settingsWindow.dgComputerAttributes.Items.Refresh()
    })

    #
    # Move Computer definition down clicked
    #
    $settingsWindow.btnDownComputerAttributeDefinition.add_Click({
        Move-Down -SourceList $settingsUnconfirmed.ComputerAttributeDefinitions -Item $settingsWindow.dgComputerAttributes.SelectedItem
        $settingsWindow.dgComputerAttributes.Items.Refresh()
    })

    #
    # Add User definition clicked
    #
    $settingsWindow.btnAddUserAttributeDefinition.add_Click({
        $settingsUnconfirmed.UserAttributeDefinitions.Add((New-Object AttributeDefinition))
        $settingsWindow.dgUserAttributes.Items.Refresh()
    })

    #
    # Remove User definition clicked
    #
    $settingsWindow.btnRemoveUserAttributeDefinition.add_Click({
        $settingsUnconfirmed.UserAttributeDefinitions.Remove($settingsWindow.dgUserAttributes.SelectedItem)
        $settingsWindow.dgUserAttributes.Items.Refresh()
    })

    #
    # Move User definition up clicked
    #
    $settingsWindow.btnUpUserAttributeDefinition.add_Click({
        Move-Up -SourceList $settingsUnconfirmed.UserAttributeDefinitions -Item $settingsWindow.dgUserAttributes.SelectedItem
        $settingsWindow.dgUserAttributes.Items.Refresh()
    })

    #
    # Move User definition down clicked
    #
    $settingsWindow.btnDownUserAttributeDefinition.add_Click({
        Move-Down -SourceList $settingsUnconfirmed.UserAttributeDefinitions -Item $settingsWindow.dgUserAttributes.SelectedItem
        $settingsWindow.dgUserAttributes.Items.Refresh()
    })

    #
    # Ok button clicked
    #
    $settingsWindow.btnOk.add_Click({
        Write-Log -LogString $settings.ComputerAttributeDefinitions[0].FriendlyName -Severity "Informational"
        Write-Log -LogString $settingsUnconfirmed.ComputerAttributeDefinitions[0].FriendlyName -Severity "Informational"

        if(Resolve-AttributeDefinitionsFromSchema -ComputerAttributeDefinitions $settingsUnconfirmed.ComputerAttributeDefinitions -UserAttributeDefinitions $settingsUnconfirmed.UserAttributeDefinitions)
        {
            $script:settings = $settingsUnconfirmed
            
            Write-Log -LogString ("Checking if settings folder [" + ($script:settingsFile.Substring(0, $script:settingsFile.LastIndexOf("\")) + "\") + "] exists") -Severity "Debug"
            if(!(Test-Path -Path ($script:settingsFile.Substring(0, $script:settingsFile.LastIndexOf("\")) + "\")))
            {
                Write-Log -LogString ("Settings folder did not exist, creating it...") -Severity "Debug"
                New-Item -ItemType Directory -Path ($script:settingsFile.Substring(0, $script:settingsFile.LastIndexOf("\")) + "\")
            }
            Write-Log -LogString ("Settings folder already exists") -Severity "Debug"

            #Write-Host ("Split: " + $script:settingsFile.Substring(0, $script:settingsFile.LastIndexOf("\")))
            
            
            Export-Clixml -Path $script:settingsFile -InputObject $script:settings
            $settingsWindow.Window.Close()
        }
    })

    #
    # Cancel button clicked
    #
    $settingsWindow.btnCancel.add_Click({
        $settingsWindow.Window.Close()
    })

    #Returns a copy of the Settings class.
    function Get-SettingsClone()
    {
        Param(
        [Parameter(Mandatory=$true)]
        $Source
        )

        function Get-AttributeDefinitionsClone()
        {
            Param(
            [Parameter(Mandatory=$true)]
            $Source
            )

            $Clone = New-Object System.Collections.Generic.List[AttributeDefinition]
            foreach($attrDef in $Source)
            {
                $attrDefClone = New-Object AttributeDefinition

                foreach ($Property in ($attrDef | Get-Member -MemberType Property))
                {
                    $attrDefclone.$($Property.Name) = $attrDef.$($Property.Name)
                }
                $Clone.Add($attrDefClone)
            }
            return $Clone
        }

        $Clone = New-Object Settings

        foreach ($Property in ($Source | Get-Member -MemberType Property))
        {
            if($Property.Name -eq "ComputerAttributeDefinitions" -or $Property.Name -eq "UserAttributeDefinitions")
            {
                $Clone.$($Property.Name) = Get-AttributeDefinitionsClone -Source $Source.$($Property.Name)
            }
            else
            {
                $Clone.$($Property.Name) = $Source.$($Property.Name)
            }

        }
        return $Clone
    }

    $settingsUnconfirmed = Get-SettingsClone -Source $script:settings
    $settingsWindow.Window.DataContext = $settingsUnconfirmed

    $settingsWindow.Window.ShowDialog() | Out-Null
}
###############################################################################
#.SYNOPSIS
#    Loads ADComputer and ADUser objects from Active Directory and
#    sets resulting collection as ItemsSource for the DataGrids
#
#.PARAMETER SearchBase
#    SearchBase can be either ADOrganizationalUnit, ADGroup or ADDomain
#
#.PARAMETER SearchString
#    Will filter results based on SearchString
#
#.PARAMETER SearchScope
#    SubTree or OneLevel
#
###############################################################################
function Update-DataGridsItemsSources
{
    Param
    (
        [Parameter(Mandatory=$false)]
        [Object]$SearchBase = (Get-SearchBase),
        [Parameter(Mandatory=$false)]
        [string]$SearchString = $null,
        [Parameter(Mandatory=$false)]
        [string]$SearchScope = (Get-SearchScope)
    )

    #
    # Generate -LDAPFilter string. This is more flexible then using regular -Filter
    #
    if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
    {
        $computerLdapFilter = "(&(objectCategory=computer)"
        $userLdapFilter = "(&(objectCategory=person)"
    }
    elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
    {
        if($SearchScope -eq "Subtree")
        {
            $computerLdapFilter = "(&(objectCategory=computer)(memberof:1.2.840.113556.1.4.1941:=" + $SearchBase.DistinguishedName + ")"
            $userLdapFilter = "(&(objectCategory=person)(memberof:1.2.840.113556.1.4.1941:=" + $SearchBase.DistinguishedName + ")"
        }
        else
        {
            $computerLdapFilter = "(&(objectCategory=computer)(memberof=" + $SearchBase.DistinguishedName + ")"
            $userLdapFilter = "(&(objectCategory=person)(memberof=" + $SearchBase.DistinguishedName + ")"
        }
    }
    else
    {
        Write-Log -LogString "Error, unknown type as SearchBase!" -Severity "Error"
        return
    }

    #
    # If user has provided a filter string, continue generation of the -LDAPFilter string
    #
    if(![string]::IsNullOrWhiteSpace($SearchString))
    {
        $computerLdapFilter = $computerLdapFilter + "(|(cn=*" + $SearchString + "*)"
        $userLdapFilter = $userLdapFilter + "(|(cn=*" + $SearchString + "*)"
        foreach($attr in $script:settings.ComputerAttributeDefinitions)
        {
            $computerLdapFilter = ($computerLdapFilter + "(" + $attr.Attribute + "=*" + $SearchString + "*)")
        }
        foreach($attr in $script:settings.UserAttributeDefinitions)
        {
            $userLdapFilter = ($userLdapFilter + "(" + $attr.Attribute + "=*" + $SearchString + "*)")
        }
        $computerLdapFilter = $computerLdapFilter + ")"
        $userLdapFilter = $userLdapFilter + ")"
    }

    $computerLdapFilter = $computerLdapFilter + ")"
    $userLdapFilter = $userLdapFilter + ")"


    #
    # Fetch computer objects.
    #
    $startTime = Get-Date
    try
    {
        Write-Log -LogString ("Loading computers from " + $SearchBase.GetType()) -Severity "Informational"
        if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
        {
            $script:MainWindow.computersDataGrid.ItemsSource = @(Get-ADComputer -LDAPFilter $computerLdapFilter -SearchScope $SearchScope -SearchBase $SearchBase.DistinguishedName -Properties $script:settings.ComputerAttributeDefinitions.Attribute)
        }
        elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
        {
            $script:MainWindow.computersDataGrid.ItemsSource = @(Get-ADComputer -LDAPFilter $computerLdapFilter -Properties $script:settings.ComputerAttributeDefinitions.Attribute)
        }
        Write-Log -LogString ($script:MainWindow.computersDataGrid.ItemsSource.Count.ToString() + " found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error')
    }

    #
    # Fetch user objects.
    #
    $startTime = Get-Date
    try
    {
        Write-Log -LogString ("Loading users from " + $SearchBase.GetType()) -Severity "Informational"
        [System.Collections.ObjectModel.ObservableCollection[Object]]$usersCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
        {
            $script:MainWindow.usersDataGrid.ItemsSource = @(Get-ADUser -LDAPFilter $userLdapFilter -SearchScope $SearchScope -SearchBase $SearchBase.DistinguishedName -Properties $script:settings.UserAttributeDefinitions.Attribute)
        }
        elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
        {
            $script:MainWindow.usersDataGrid.ItemsSource = @(Get-ADUser -LDAPFilter $userLdapFilter -Properties $script:settings.UserAttributeDefinitions.Attribute)
        }
        Write-Log -LogString ($script:MainWindow.usersDataGrid.ItemsSource.Count.ToString() + " found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error')
    }

    #
    # Select the TabItem containing the highest items count
    #
    if($script:MainWindow.computersDataGrid.ItemsSource.Count -gt $script:MainWindow.usersDataGrid.ItemsSource.Count)
    {
        $script:MainWindow.tabItemComputers.IsSelected = $true
    }
    elseif($script:MainWindow.usersDataGrid.ItemsSource.Count -gt $script:MainWindow.computersDataGrid.ItemsSource.Count)
    {
        $script:MainWindow.tabItemUsers.IsSelected = $true
    }

    #Update some UI controls...
    Update-UI
    Update-Statistics
}
###############################################################################
#.SYNOPSIS
#    Currently not much here. Updates the TabItems header with counts of all/active/passive objects.
#
###############################################################################
function Update-Statistics
{
    $computerLimit = (Get-Date).AddDays(-$script:settings.ComputerInactiveLimit).ToFileTimeUtc()
    $userLimit = (Get-Date).AddDays(-$script:settings.UserInactiveLimit).ToFileTimeUtc()

    if($script:MainWindow.computersDataGrid.ItemsSource.Count -gt 0)
    {
        $script:MainWindow.txtTotalComputerCount.Text = @($script:MainWindow.computersDataGrid.ItemsSource).Count.ToString()
        $script:MainWindow.txtActiveComputerCount.Text = @($script:MainWindow.computersDataGrid.ItemsSource | Where-Object {$_.lastLogonTimeStamp -gt $computerLimit }).Count.ToString()
        $script:MainWindow.txtPassiveComputerCount.Text = @($script:MainWindow.computersDataGrid.ItemsSource | Where-Object {$_.lastLogonTimeStamp -lt $computerLimit }).Count.ToString()
    }
    else
    {
        $script:MainWindow.txtTotalComputerCount.Text = "0"
        $script:MainWindow.txtActiveComputerCount.Text = "0"
        $script:MainWindow.txtPassiveComputerCount.Text = "0"
    }

    if($script:MainWindow.usersDataGrid.ItemsSource.Count -gt 0)
    {
        $script:MainWindow.txtTotalUserCount.Text = @($script:MainWindow.usersDataGrid.ItemsSource).Count.ToString()
        $script:MainWindow.txtActiveUserCount.Text = @($script:MainWindow.usersDataGrid.ItemsSource | Where-Object {$_.lastLogonTimeStamp -gt $userLimit }).Count.ToString()
        $script:MainWindow.txtPassiveUserCount.Text = @($script:MainWindow.usersDataGrid.ItemsSource | Where-Object {$_.lastLogonTimeStamp -lt $userLimit }).Count.ToString()
    }
    else
    {
        $script:MainWindow.txtTotalUserCount.Text = "0"
        $script:MainWindow.txtActiveUserCount.Text = "0"
        $script:MainWindow.txtPassiveUserCount.Text = "0"
    }
}
###############################################################################
#.SYNOPSIS
#    Updates UI elements after items have been fetched from AD
#
#.DESCRIPTION
#    Updates UI Elements.
#    1. Generates DataGridColumns and sets bindings
#    2. Generates Stackpanels with subcontrols for DetailsPane
#    3. Generates Stackpanels with subcontrols for EditablAttributesPane
#
#.NOTES
#
###############################################################################
Function Update-UI
{
    #Remove all columns
    $script:MainWindow.computersDataGrid.Columns.Clear()
    $script:MainWindow.usersDataGrid.Columns.Clear()


    foreach($dg in @($script:MainWindow.computersDataGrid, $script:MainWindow.usersDataGrid))
    {
        if($dg.Name -eq "computersDataGrid")
        {
            $stkEditableAttributes = $script:MainWindow.stkEditableComputerAttributes
            $stkDetailsPane = $script:MainWindow.stkComputerDetailsPane
            $attributes = $script:settings.ComputerAttributeDefinitions
        }
        elseif($dg.Name -eq "usersDataGrid")
        {
            $stkEditableAttributes = $script:MainWindow.stkEditableUserAttributes
            $stkDetailsPane = $script:MainWindow.stkUserDetailsPane
            $attributes = $script:settings.UserAttributeDefinitions
        }

        #
        # Setup DataGrid Columns
        #
        foreach($attr in ($attributes | Where-Object {$_.DisplayIn -eq "DataGrid"}))
        {
            $dgColumn = New-Object System.Windows.Controls.DataGridTextColumn

            Write-Log -LogString ("Adding binding: "  + $attr.Attribute) -Severity "Debug"
            $dgColumn.Binding = New-Object System.Windows.Data.Binding($attr.Attribute)
            $dgColumn.Binding.Mode = [System.Windows.Data.BindingMode]::OneWay
            $dgColumn.Binding.Converter = Get-Converter -AttributeDefinition $attr

            #
            # Create tooltip
            #
            $headerGrid = New-Object System.Windows.Controls.Grid
            $headerStk = New-Object System.Windows.Controls.StackPanel
            $headerStk.Orientation = "Horizontal"
            $headerTxt = New-Object System.Windows.Controls.TextBlock
            $headerTxt.Text = $attr.FriendlyName
            $headerStk.Children.Add($headerTxt)

            if($dgColumn.Binding.Converter -ne $null)
            {
                $headerGrid.ToolTip = ("Attribute [" + $attr.Attribute + "] Converter [" +  $dgColumn.Binding.Converter.GetType().ToString() + "]")
                $path = New-Object System.Windows.Shapes.Path
                $path.Data = $script:MainWindow.Window.FindResource("infoIcon")
                $path.Fill = $script:MainWindow.Window.FindResource("iconColor")
                $path.Margin = "2"
                $path.Stretch = "Fill"
                $path.Height = "12"
                $path.Width = "12"
                $headerStk.Children.Add($path)
            }
            else
            {
                $headerGrid.ToolTip = ("Attribute [" + $attr.Attribute + "]")
            }

            $headerGrid.Children.Add($headerStk)
            $dgColumn.Header = $headerGrid

            # Need to check if converter is [ADPropertyValueCollectionConverter]. If it is, sorting will not work without setting sortmemberpath to index in list.
            if($dgColumn.Binding.Converter -ne $null -and $dgColumn.Binding.Converter.GetType() -eq [ADPropertyValueCollectionConverter])
            {
                $dgColumn.SortMemberPath = ($attr.attribute + "[0]")
            }

            $dg.Columns.Add($dgColumn)
        }

        #
        # Add items to details pane
        #
        $stkDetailsPane.Children.Clear()
        foreach($attr in ($attributes | Where-Object {$_.DisplayIn -eq "DetailsPane"}))
        {
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $stackPanel.Orientation = "Horizontal"

            $txtFriendlyName = New-Object System.Windows.Controls.TextBlock
            $txtFriendlyName.Text = $attr.FriendlyName
            $txtFriendlyName.MinWidth = "200"

            $txtValue = New-Object System.Windows.Controls.TextBlock
            $txtValue.DataContext = $dg
            $binding = New-Object System.Windows.Data.Binding("SelectedItem." + $attr.Attribute)
            $binding.Mode = [System.Windows.Data.BindingMode]::OneWay
            $binding.Converter = Get-Converter -AttributeDefinition $attr
            [void][System.Windows.Data.BindingOperations]::SetBinding($txtValue,[System.Windows.Controls.TextBlock]::TextProperty, $binding)
            $stackPanel.Children.Add($txtFriendlyName) | Out-Null
            $stackPanel.Children.Add($txtValue) | Out-Null
            $stkDetailsPane.Children.Add($stackPanel)
        }

        #
        # Add items to editable attributes pane
        #
        $stkEditableAttributes.Children.Clear()
        foreach($attr in ($attributes | Where-Object {$_.IsEditable -eq $true}))
        {
            if($attr.IsEditable)
            {
                $stackPanel = New-Object System.Windows.Controls.StackPanel
                $stackPanel.Orientation = "Horizontal"
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = $attr.FriendlyName
                $stackPanel.Children.Add($textBlock) | Out-Null
                if($attr.IsSingleValued)
                {
                    $stackPanel.ToolTip = "Attribute: " + $attr.Attribute
                }
                else
                {
                    $path = New-Object System.Windows.Shapes.Path
                    $path.Data = $script:MainWindow.Window.FindResource("infoIcon")
                    $path.Fill = $script:MainWindow.Window.FindResource("iconColor")
                    $path.Margin = "2"
                    $path.Stretch = "Fill"
                    $path.Height = "12"
                    $path.Width = "12"
                    $stackPanel.Children.Add($path) | Out-Null
                    $stackPanel.ToolTip = "Attribute: " + $attr.Attribute + " [This is a multi value attribute, use '" + $script:ADPropertyValueCollectionConverter.Separator + "' as separator]"
                }

                $stkEditableAttributes.Children.Add($stackPanel) | Out-Null
                $comboBox = New-Object System.Windows.Controls.ComboBox
                $comboBox.Margin = "0,0,0,5"
                $comboBox.IsEditable = $true

                #
                # Set Combobox Tag. This is used when setting new values (to connect comboboxes to a specific attribute)
                #
                $comboBox.Tag = $attr.Attribute
                if($dg.ItemsSource) #Needs to check for $null here. Set-StrictMode will complain otherwise...
                {
                    #
                    # Need to filter out objects missing the $attr.Attribute property. Set-StrictMode will complain otherwise...
                    #
                    $lst = @($dg.ItemsSource | Where-Object {Get-Member -InputObject $_ -Name $attr.Attribute -Membertype Properties})
                    if($lst.Count -gt 0)
                    {
                        $comboBox.ItemsSource = @($lst.($attr.Attribute) | Where-Object {$_ -ne $null} | Sort-Object | Get-Unique )
                    }
                }

                $binding = New-Object System.Windows.Data.Binding("SelectedItem." + $attr.Attribute)
                $binding.ElementName = $dg.Name
                $binding.Converter = Get-Converter -AttributeDefinition $attr
                $binding.Mode = [System.Windows.Data.BindingMode]::OneWay
                [void][System.Windows.Data.BindingOperations]::SetBinding($comboBox,[System.Windows.Controls.ComboBox]::TextProperty, $binding)
                $stkEditableAttributes.Children.Add($comboBox) | Out-Null
            }
        }
    }
}
###############################################################################
#.SYNOPSIS
#    Custom function to get function name prepended to debug output
#
#.PARAMETER DebugString
#    String to be written as debug output
#
#.EXAMPLE
#    Write-ADIDebug "DebugString"
#
###############################################################################
function Write-ADIDebug
{
    Param(
        [Parameter(Mandatory=$false)]
            $DebugString
        )

        $DebugString = "[" + $((Get-PSCallStack)[1].Command) + "]: " + $DebugString
        Write-Debug $DebugString
}
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
#endregion

#region Classes
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
    [bool] $ShowVerboseOutput = $true
    [bool] $ShowDebugOutput = $true   
    [int] $ComputerInactiveLimit = 90
    [int] $UserInactiveLimit = 90
    [string] $IconColor = "#336699"
}
#endregion

#region XAML
[xml]$script:xamlMainWindow = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window" WindowStartupLocation = "CenterScreen"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Width = "1366" Height = "800" ShowInTaskbar = "True">
   <Window.Resources>
        <Geometry x:Key="expandIcon">F1 M 19,19L 57,19L 57,57L 19,57L 19,19 Z M 54,54L 54,22.0001L 22,22.0001L 22,54L 54,54 Z M 41,24.0001L 41,52L 24,52L 24,40L 33.6666,40L 30,45L 34,45L 39,38L 34,31L 30,31L 33.6666,36L 24,36L 24,24L 41,24.0001 Z</Geometry>
        <Geometry x:Key="collapseIcon">F1 M 57,19L 19,19L 19,57L 57,57L 57,19 Z M 22,54L 22,22.0001L 54,22.0001L 54,54L 22,54 Z M 24,24.0001L 24,52L 41,52L 40.9999,40L 31.3333,40L 35,45L 31,45L 26,38L 31,31L 35,31L 31.3333,36L 40.9999,36L 41,24L 24,24.0001 Z</Geometry>
        <Geometry x:Key="exportIcon">F1 M 42,24L 57,24L 57,52L 42,52L 42,50L 47,50L 47,46L 42,46L 42,45L 47,45L 47,41L 42,41L 42,40L 47,40L 47,36L 42,36L 42,35L 47,35L 47,31L 42,31L 42,30L 47,30L 47,26L 42,26L 42,24 Z M 54.9995,50.0005L 54.9997,46.0003L 47.9995,46.0003L 47.9995,50.0005L 54.9995,50.0005 Z M 54.9996,41.0004L 47.9995,41.0004L 47.9995,45.0003L 54.9997,45.0003L 54.9996,41.0004 Z M 54.9996,36.0004L 47.9995,36.0004L 47.9995,40.0004L 54.9996,40.0004L 54.9996,36.0004 Z M 54.9996,31.0004L 47.9995,31.0004L 47.9995,35.0004L 54.9996,35.0004L 54.9996,31.0004 Z M 54.9995,26.0005L 47.9995,26.0005L 47.9995,30.0004L 54.9996,30.0004L 54.9995,26.0005 Z M 18.9997,23.7503L 40.9994,19.7506L 40.9994,56.2506L 18.9997,52.2503L 18.9997,23.7503 Z M 34.6404,44.5147L 31.3367,37.4084L 34.5522,30.4699L 31.9399,30.5805L 30.2234,34.6963L 30.0162,35.3903L 29.8872,35.8892L 29.8596,35.8895C 29.4574,34.1248 28.7481,32.4436 28.1318,30.7417L 25.2803,30.8624L 28.2549,37.4637L 24.997,44.0621L 27.7904,44.1932L 29.5296,39.8757L 29.7578,38.9297L 29.7876,38.93C 30.2317,40.8236 31.1236,42.5844 31.861,44.3843L 34.6404,44.5147 Z</Geometry>
        <Geometry x:Key="computerIcon">M6,2C4.89,2 4,2.89 4,4V12C4,13.11 4.89,14 6,14H18C19.11,14 20,13.11 20,12V4C20,2.89 19.11,2 18,2H6M6,4H18V12H6V4M4,15C2.89,15 2,15.89 2,17V20C2,21.11 2.89,22 4,22H20C21.11,22 22,21.11 22,20V17C22,15.89 21.11,15 20,15H4M8,17H20V20H8V17M9,17.75V19.25H13V17.75H9M15,17.75V19.25H19V17.75H15Z</Geometry>
        <Geometry x:Key="userIcon">M12,4A4,4 0 0,1 16,8A4,4 0 0,1 12,12A4,4 0 0,1 8,8A4,4 0 0,1 12,4M12,14C16.42,14 20,15.79 20,18V20H4V18C4,15.79 7.58,14 12,14Z</Geometry>
        <Geometry x:Key="enabledIcon">M4,11V13H16L10.5,18.5L11.92,19.92L19.84,12L11.92,4.08L10.5,5.5L16,11H4Z</Geometry>
        <Geometry x:Key="disabledIcon">M11,4H13V16L18.5,10.5L19.92,11.92L12,19.84L4.08,11.92L5.5,10.5L11,16V4Z</Geometry>
        <Geometry x:Key="arrowsIcon">M13,11H18L16.5,9.5L17.92,8.08L21.84,12L17.92,15.92L16.5,14.5L18,13H13V18L14.5,16.5L15.92,17.92L12,21.84L8.08,17.92L9.5,16.5L11,18V13H6L7.5,14.5L6.08,15.92L2.16,12L6.08,8.08L7.5,9.5L6,11H11V6L9.5,7.5L8.08,6.08L12,2.16L15.92,6.08L14.5,7.5L13,6V11Z</Geometry>
        <Geometry x:Key="folderIcon">M10,4H4C2.89,4 2,4.89 2,6V18A2,2 0 0,0 4,20H20A2,2 0 0,0 22,18V8C22,6.89 21.1,6 20,6H12L10,4Z</Geometry>
        <Geometry x:Key="openFolderIcon">M19,20H4C2.89,20 2,19.1 2,18V6C2,4.89 2.89,4 4,4H10L12,6H19A2,2 0 0,1 21,8H21L4,8V18L6.14,10H23.21L20.93,18.5C20.7,19.37 19.92,20 19,20Z</Geometry>
        <Geometry x:Key="searchIcon">M9.5,3A6.5,6.5 0 0,1 16,9.5C16,11.11 15.41,12.59 14.44,13.73L14.71,14H15.5L20.5,19L19,20.5L14,15.5V14.71L13.73,14.44C12.59,15.41 11.11,16 9.5,16A6.5,6.5 0 0,1 3,9.5A6.5,6.5 0 0,1 9.5,3M9.5,5C7,5 5,7 5,9.5C5,12 7,14 9.5,14C12,14 14,12 14,9.5C14,7 12,5 9.5,5Z</Geometry>
        <Geometry x:Key="treeIcon">M3,3H9V7H3V3M15,10H21V14H15V10M15,17H21V21H15V17M13,13H7V18H13V20H7L5,20V9H7V11H13V13Z</Geometry>
        <Geometry x:Key="editIcon">M21.7,13.35L20.7,14.35L18.65,12.3L19.65,11.3C19.86,11.08 20.21,11.08 20.42,11.3L21.7,12.58C21.92,12.79 21.92,13.14 21.7,13.35M12,18.94L18.07,12.88L20.12,14.93L14.06,21H12V18.94M4,2H18A2,2 0 0,1 20,4V8.17L16.17,12H12V16.17L10.17,18H4A2,2 0 0,1 2,16V4A2,2 0 0,1 4,2M4,6V10H10V6H4M12,6V10H18V6H12M4,12V16H10V12H4Z</Geometry>
        <Geometry x:Key="consoleIcon">M20,19V7H4V19H20M20,3A2,2 0 0,1 22,5V19A2,2 0 0,1 20,21H4A2,2 0 0,1 2,19V5C2,3.89 2.9,3 4,3H20M13,17V15H18V17H13M9.58,13L5.57,9H8.4L11.7,12.3C12.09,12.69 12.09,13.33 11.7,13.72L8.42,17H5.59L9.58,13Z</Geometry>
        <Geometry x:Key="overviewIcon">M21,11H13V3A8,8 0 0,1 21,11M19,13C19,15.78 17.58,18.23 15.43,19.67L11.58,13H19M11,21C8.22,21 5.77,19.58 4.33,17.43L10.82,13.68L14.56,20.17C13.5,20.7 12.28,21 11,21M3,13A8,8 0 0,1 11,5V12.42L3.83,16.56C3.3,15.5 3,14.28 3,13Z</Geometry>
        <Geometry x:Key="reloadIcon">M19,12H22.32L17.37,16.95L12.42,12H16.97C17,10.46 16.42,8.93 15.24,7.75C12.9,5.41 9.1,5.41 6.76,7.75C4.42,10.09 4.42,13.9 6.76,16.24C8.6,18.08 11.36,18.47 13.58,17.41L15.05,18.88C12,20.69 8,20.29 5.34,17.65C2.22,14.53 2.23,9.47 5.35,6.35C8.5,3.22 13.53,3.21 16.66,6.34C18.22,7.9 19,9.95 19,12Z</Geometry>
        <Geometry x:Key="groupIcon">M8,8V12H13V8H8M1,1H5V2H19V1H23V5H22V19H23V23H19V22H5V23H1V19H2V5H1V1M5,19V20H19V19H20V5H19V4H5V5H4V19H5M6,6H15V10H18V18H8V14H6V6M15,14H10V16H16V12H15V14Z</Geometry>
        <Geometry x:Key="sourceIcon">M18,15H16V17H18M18,11H16V13H18M20,19H12V17H14V15H12V13H14V11H12V9H20M10,7H8V5H10M10,11H8V9H10M10,15H8V13H10M10,19H8V17H10M6,7H4V5H6M6,11H4V9H6M6,15H4V13H6M6,19H4V17H6M12,7V3H2V21H22V7H12Z</Geometry>
        <Geometry x:Key="settingsIcon">M12,15.5A3.5,3.5 0 0,1 8.5,12A3.5,3.5 0 0,1 12,8.5A3.5,3.5 0 0,1 15.5,12A3.5,3.5 0 0,1 12,15.5M19.43,12.97C19.47,12.65 19.5,12.33 19.5,12C19.5,11.67 19.47,11.34 19.43,11L21.54,9.37C21.73,9.22 21.78,8.95 21.66,8.73L19.66,5.27C19.54,5.05 19.27,4.96 19.05,5.05L16.56,6.05C16.04,5.66 15.5,5.32 14.87,5.07L14.5,2.42C14.46,2.18 14.25,2 14,2H10C9.75,2 9.54,2.18 9.5,2.42L9.13,5.07C8.5,5.32 7.96,5.66 7.44,6.05L4.95,5.05C4.73,4.96 4.46,5.05 4.34,5.27L2.34,8.73C2.21,8.95 2.27,9.22 2.46,9.37L4.57,11C4.53,11.34 4.5,11.67 4.5,12C4.5,12.33 4.53,12.65 4.57,12.97L2.46,14.63C2.27,14.78 2.21,15.05 2.34,15.27L4.34,18.73C4.46,18.95 4.73,19.03 4.95,18.95L7.44,17.94C7.96,18.34 8.5,18.68 9.13,18.93L9.5,21.58C9.54,21.82 9.75,22 10,22H14C14.25,22 14.46,21.82 14.5,21.58L14.87,18.93C15.5,18.67 16.04,18.34 16.56,17.94L19.05,18.95C19.27,19.03 19.54,18.95 19.66,18.73L21.66,15.27C21.78,15.05 21.73,14.78 21.54,14.63L19.43,12.97Z</Geometry>
        <Geometry x:Key="infoIcon">M11,9H13V7H11M12,20C7.59,20 4,16.41 4,12C4,7.59 7.59,4 12,4C16.41,4 20,7.59 20,12C20,16.41 16.41,20 12,20M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M11,17H13V11H11V17Z</Geometry>
        <Geometry x:Key="lapsIcon">F1 M 15.8333,25.3333L 60.1667,25.3333L 60.1667,52.25L 15.8333,52.25L 15.8333,25.3333 Z M 19,28.5L 19,49.0833L 57,49.0833L 57,28.5L 19,28.5 Z M 32.015,39.1558L 28.591,39.5833L 31.0017,42.1286L 28.7652,43.6367L 27.0829,40.4938L 25.3967,43.6367L 23.1483,42.1286L 25.5708,39.5833L 22.135,39.1558L 23.0256,36.7967L 26.1883,38.0633L 25.6817,34.5167L 28.4683,34.5167L 27.9617,38.0633L 31.1006,36.7967L 32.015,39.1558 Z M 43.6683,39.1558L 40.2444,39.5833L 42.655,42.1285L 40.4185,43.6367L 38.7362,40.4938L 37.05,43.6367L 34.8017,42.1285L 37.2242,39.5833L 33.7883,39.1558L 34.679,36.7967L 37.8417,38.0633L 37.335,34.5167L 40.1217,34.5167L 39.615,38.0633L 42.754,36.7967L 43.6683,39.1558 Z M 45.5208,47.5L 45.5208,30.0833L 47.5,30.0833L 47.5,47.5L 45.5208,47.5 Z</Geometry>
        <Geometry x:Key="debugIcon">F1 M 46.5,19C 47.8807,19 49,20.1193 49,21.5C 49,22.8807 47.8807,24 46.5,24L 45.8641,23.9184L 43.5566,26.8718C 45.1489,28.0176 46.5309,29.6405 47.6023,31.6025C 44.8701,32.4842 41.563,33 38,33C 34.4369,33 31.1299,32.4842 28.3977,31.6025C 29.4333,29.7061 30.7591,28.1265 32.2844,26.9882L 29.9221,23.9646C 29.7849,23.9879 29.6438,24 29.5,24C 28.1193,24 27,22.8808 27,21.5C 27,20.1193 28.1193,19 29.5,19C 30.8807,19 32,20.1193 32,21.5C 32,22.0018 31.8521,22.4691 31.5976,22.8607L 34.0019,25.938C 35.2525,25.3305 36.5982,25 38,25C 39.3339,25 40.617,25.2993 41.8156,25.8516L 44.2947,22.6786C 44.1066,22.3274 44,21.9262 44,21.5C 44,20.1193 45.1193,19 46.5,19 Z M 54.5,40C 55.3284,40 56,40.6716 56,41.5C 56,42.3284 55.3284,43 54.5,43L 49.9511,43C 49.88,44.0847 49.7325,45.1391 49.5162,46.1531L 54.8059,48.6197C 55.5567,48.9698 55.8815,49.8623 55.5314,50.6131C 55.1813,51.3639 54.2889,51.6887 53.5381,51.3386L 48.6665,49.067C 46.8161,53.9883 43.2172,57.4651 39,57.9435L 39,34.9864C 42.541,34.8897 45.7913,34.283 48.4239,33.3201L 48.6187,33.8074L 53.73,31.8454C 54.5034,31.5485 55.371,31.9348 55.6679,32.7082C 55.9648,33.4816 55.5785,34.3492 54.8051,34.6461L 49.482,36.6895C 49.717,37.7515 49.8763,38.859 49.9511,40L 54.5,40 Z M 21.5,40L 26.0489,40C 26.1237,38.859 26.2829,37.7516 26.518,36.6895L 21.1949,34.6461C 20.4215,34.3492 20.0352,33.4816 20.332,32.7082C 20.6289,31.9348 21.4966,31.5485 22.27,31.8454L 27.3812,33.8074L 27.5761,33.3201C 30.2087,34.283 33.4589,34.8897 37,34.9864L 37,57.9435C 32.7827,57.4651 29.1838,53.9883 27.3335,49.067L 22.4618,51.3387C 21.711,51.6888 20.8186,51.3639 20.4685,50.6131C 20.1184,49.8623 20.4432,48.9699 21.194,48.6198L 26.4838,46.1531C 26.2674,45.1392 26.12,44.0847 26.0489,43L 21.5,43C 20.6716,43 20,42.3285 20,41.5C 20,40.6716 20.6716,40 21.5,40 Z</Geometry>
        <SolidColorBrush x:Key="iconColor">#336699</SolidColorBrush>

        <sys:Double x:Key="FontSize">13</sys:Double>
        <sys:Double x:Key="ButtonSize">28</sys:Double>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
         <Style TargetType="DataGridRow">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>

        <Style TargetType="{x:Type TreeViewItem}">
            <Setter Property="Foreground" Value="Black"/>
            <Setter Property="HeaderTemplate">
                <Setter.Value>
                    <DataTemplate>
                        <StackPanel Orientation="Horizontal">
                            <Path Name="pathExpandCollapse" Margin="1" Stretch="Uniform" Fill="{StaticResource iconColor}">
                                <Path.Style>
                                    <Style TargetType="Path">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource AncestorType=TreeViewItem}, Path=IsExpanded}" Value="True">
                                                <Setter Property="Data" Value="{StaticResource openFolderIcon}"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource AncestorType=TreeViewItem}, Path=IsExpanded}" Value="False">
                                                <Setter Property="Data" Value="{StaticResource folderIcon}"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </Path.Style>
                            </Path>
                            <TextBlock Text="{Binding}" FontSize="{StaticResource FontSize}" Margin="5,0" VerticalAlignment="Center" />
                        </StackPanel>
                    </DataTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition>
                <!--Sidebar-->
                <ColumnDefinition.Style>
                    <Style TargetType="ColumnDefinition">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="True">
                                <Setter Property="MaxWidth" Value="0"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="False">
                                <Setter Property="Width" Value="Auto"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </ColumnDefinition.Style>
            </ColumnDefinition>

            <ColumnDefinition>
                <!--Gridsplitter-->
                <ColumnDefinition.Style>
                    <Style TargetType="ColumnDefinition">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="True">
                                <Setter Property="MaxWidth" Value="0"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="False">
                                <Setter Property="Width" Value="5"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </ColumnDefinition.Style>
            </ColumnDefinition>

            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Grid Grid.Row="1" Grid.Column="0" Name="grdSidebar" MinWidth="350" Margin="2,0,4,0">
            <TabControl>
                <TabItem>
                    <TabItem.Header>
                        <StackPanel Orientation="Horizontal">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource sourceIcon}"/>
                            <TextBlock VerticalAlignment="Center" Text="Source" Margin="4,0,0,0"/>
                        </StackPanel>
                    </TabItem.Header>

                    <TabControl>
                        <TabItem>
                            <TabItem.Header>
                                <StackPanel Orientation="Horizontal">
                                    <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource treeIcon}"/>
                                    <TextBlock VerticalAlignment="Center" Text="Organizational Units" Margin="4,0,0,0"/>
                                </StackPanel>
                            </TabItem.Header>
                            <Grid>

                               <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>

                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Grid.Column="0" Margin="0,0,0,2" Name="txtOUFilter"/>

                                    <Button Grid.Column="1" Height="25" Width="25" Name="btnFilterOUs" ToolTip="Search Organizational Units" Background="Transparent" Margin="2,0,0,2">
                                        <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource searchIcon}"/>
                                    </Button>

                                    <Button Grid.Column="2" Height="25" Width="25" Name="btnReloadOUs" ToolTip="Reload Organizational Units" Background="Transparent" Margin="2,0,0,2">
                                        <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource reloadIcon}"/>
                                    </Button>

                                    <CheckBox Grid.Column="3" Name="chkRecursiveOUSearch" IsChecked="True" VerticalContentAlignment="Center" VerticalAlignment="Center" Content="Recurse" Margin="2,0,0,2"/>
                                </Grid>

                                <Separator Grid.Row="1" Margin="-2,0,-2,0"/>

                                <TreeView Grid.Row="2" Name="tvOUBrowser" BorderThickness="0" Margin="0,2,0,0"/>
                            </Grid>
                        </TabItem>

                        <TabItem>
                            <TabItem.Header>
                                <StackPanel Orientation="Horizontal">
                                    <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource groupIcon}"/>
                                    <TextBlock VerticalAlignment="Center" Text="Groups" Margin="4,0,0,0"/>
                                </StackPanel>
                            </TabItem.Header>
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>

                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>

                                    <TextBox Grid.Column="0" Margin="0,0,0,2" Name="txtGroupFilter"/>

                                    <Button Grid.Column="1" Height="25" Width="25" Name="btnFilterGroups" ToolTip="Search groups" Background="Transparent" Margin="2,0,0,2">
                                        <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource searchIcon}"/>
                                    </Button>

                                    <Button Grid.Column="2" Height="25" Width="25" Name="btnReloadGroups" ToolTip="Reload groups" Background="Transparent" Margin="2,0,0,2">
                                        <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource reloadIcon}"/>
                                    </Button>

                                    <CheckBox Grid.Column="3" Name="chkRecursiveGroupSearch" IsChecked="True" VerticalContentAlignment="Center" VerticalAlignment="Center" Content="Recurse" Margin="2,0,0,2"/>
                                </Grid>

                                <Separator Grid.Row="1" Margin="-2,0,-2,0"/>

                                <ListBox Grid.Row="2" Name="lstBoxGroups" BorderThickness="0" Margin="0,2,0,0">
                                    <ListBox.ItemTemplate>
                                        <DataTemplate>
                                            <StackPanel Orientation="Horizontal">
                                                <Path Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource groupIcon}"/>
                                                <TextBlock VerticalAlignment="Center" Text="{Binding Path=Name}" Margin="5,0"/>                                                
                                            </StackPanel>
                                        </DataTemplate>
                                    </ListBox.ItemTemplate>
                                </ListBox>

                            </Grid>
                        </TabItem>
                    </TabControl>
                </TabItem>
                <TabItem>
                    <TabItem.Header>
                        <StackPanel Orientation="Horizontal">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource editIcon}"/>
                            <TextBlock VerticalAlignment="Center" Text="Edit" Margin="4,0,0,0"/>
                        </StackPanel>
                    </TabItem.Header>
                    <Grid>
                        <ScrollViewer HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                            <Grid>
                                <Grid>
                                    <Grid.Style>
                                        <Style TargetType="Grid">
                                            <Setter Property="IsEnabled" Value="True"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding ElementName=tabItemComputers, Path=IsSelected}" Value="True">
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding ElementName=tabItemComputers, Path=IsSelected}" Value="False">
                                                    <Setter Property="Visibility" Value="Collapsed"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding ElementName=computersDataGrid, Path=SelectedIndex}" Value="-1">
                                                    <Setter Property="IsEnabled" Value="False"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </Grid.Style>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <StackPanel Margin="2,5,0,5" Grid.Row="0" Orientation="Horizontal">
                                        <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource computerIcon}"/>
                                        <TextBlock VerticalAlignment="Center" Text="{Binding ElementName=computersDataGrid, Path=SelectedItem.Name}" Margin="4,0,0,0"/>
                                    </StackPanel>

                                    <Expander Grid.Row="1" IsExpanded="True" Header="Editable attributes">
                                        <Grid Margin="5,5,5,0">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <StackPanel Grid.Row="0" Orientation="Vertical" Name="stkEditableComputerAttributes"/>
                                            <Button Grid.Row="2" Name="btnCommitComputerChanges" Content="Commit changes" Background="Transparent"/>
                                        </Grid>
                                    </Expander>

                                    <Expander Grid.Row="2" Margin="0,5,0,0" IsExpanded="True" Header="Read-Only attributes">
                                        <Grid Margin="0,5,0,0">
                                            <StackPanel Name="stkComputerDetailsPane" Orientation="Vertical"/>
                                        </Grid>
                                    </Expander>
                                </Grid>

                                <Grid>
                                    <Grid.Style>
                                        <Style TargetType="Grid">
                                            <Setter Property="IsEnabled" Value="True"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding ElementName=tabItemUsers, Path=IsSelected}" Value="True">
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding ElementName=tabItemUsers, Path=IsSelected}" Value="False">
                                                    <Setter Property="Visibility" Value="Collapsed"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding ElementName=usersDataGrid, Path=SelectedIndex}" Value="-1">
                                                    <Setter Property="IsEnabled" Value="False"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </Grid.Style>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <StackPanel Margin="2,5,0,5" Grid.Row="0" Orientation="Horizontal">
                                        <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource userIcon}"/>
                                        <TextBlock VerticalAlignment="Center" Text="{Binding ElementName=usersDataGrid, Path=SelectedItem.Name}" Margin="4,0,0,0"/>
                                    </StackPanel>

                                    <Expander Grid.Row="1" IsExpanded="True" Header="Editable attributes">
                                        <Grid Margin="0,5,0,0">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <StackPanel Grid.Row="0" Orientation="Vertical" Name="stkEditableUserAttributes"/>
                                            <Button Grid.Row="2" Name="btnCommitUserChanges" Content="Commit changes" Background="Transparent"/>
                                        </Grid>
                                    </Expander>

                                    <Expander Grid.Row="2" Margin="0,5,0,0" IsExpanded="True" Header="Read-Only attributes">
                                        <Grid Margin="0,5,0,0">
                                            <StackPanel Name="stkUserDetailsPane" Orientation="Vertical"/>
                                        </Grid>
                                    </Expander>
                                </Grid>
                            </Grid>
                        </ScrollViewer>
                    </Grid>
                </TabItem>
            </TabControl>
        </Grid>

        <GridSplitter Name="grdSplitter" Grid.Column="1" Width="5" HorizontalAlignment="Stretch"/>

        <Grid Grid.Row="0" Grid.Column="2" Name="grdButtonBar">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal">

                <ToggleButton Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" Grid.Row="0" Grid.Column="1" Name="btnExpandCollapseSideBar" Background="Transparent" Margin="0,0,0,2">
                    <Path Name="pathExpandCollapse" Margin="1" Stretch="Uniform" Fill="{StaticResource iconColor}">
                        <Path.Style>
                            <Style TargetType="Path">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="False">
                                        <Setter Property="Data" Value="{StaticResource collapseIcon}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding ElementName=btnExpandCollapseSideBar, Path=IsChecked}" Value="True">
                                        <Setter Property="Data" Value="{StaticResource expandIcon}"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </Path.Style>
                    </Path>
                </ToggleButton>

                <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="2,0"/>

                <TextBox Name="txtFilter" Width="200" VerticalContentAlignment="Center" Margin="0,0,0,2" />

                <Button Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" Name="btnFilter" ToolTip="Load / Refresh objects" Background="Transparent" Margin="2,0,0,2">
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource searchIcon}"/>
                </Button>

                <Button Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" ToolTip="Export to CSV" Name="btnExportData" Background="Transparent" Margin="2,0,0,2">
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource exportIcon}"/>
                </Button>

                <Button Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" ToolTip="Get LAPS Password" Name="btnGetLapsPassword" Background="Transparent" Margin="2,0,0,2">
                     <Button.Style>
                            <Style TargetType="Button">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding ElementName=tabItemComputers, Path=IsSelected}" Value="True">
                                        <Setter Property="Visibility" Value="Visible"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding ElementName=tabItemComputers, Path=IsSelected}" Value="False">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding ElementName=computersDataGrid, Path=SelectedIndex}" Value="-1">
                                        <Setter Property="IsEnabled" Value="False"/>
                                    </DataTrigger>                                    
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource lapsIcon}"/>
                </Button>

            </StackPanel>

            <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal">
                <Button Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" ToolTip="Show debug view" Name="btnDebugView" Background="Transparent" Margin="2,0,0,2">
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource debugIcon}"/>
                </Button>

                <Button Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" ToolTip="Settings" Name="btnSettings" Background="Transparent" Margin="2,0,0,2">
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource settingsIcon}"/>
                </Button>

                <ToggleButton Height="{StaticResource ButtonSize}" Width="{StaticResource ButtonSize}" HorizontalAlignment="Right" ToolTip="Show/Hide console" IsChecked="False" Name="btnShowHideConsole" Background="Transparent" Margin="2,0,2,2">
                    <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource consoleIcon}"/>
                </ToggleButton>
            </StackPanel>

            <!-- DATAGRID COLUMN -->
            <Grid Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TabControl Grid.Row="1" Name="tabControlDataGrids">
                    <TabItem Name="tabItemComputers">
                        <TabItem.Header>
                            <StackPanel Orientation="Horizontal">
                                <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource computerIcon}"/>
                                <TextBlock VerticalAlignment="Center" Text=" Computers"/>
                                <TextBlock VerticalAlignment="Center" Text=" ["/>
                                <TextBlock VerticalAlignment="Center" Name="txtTotalComputerCount" ToolTip="Total computers"/>
                                <TextBlock VerticalAlignment="Center" Text="/"/>
                                <TextBlock VerticalAlignment="Center" Name="txtActiveComputerCount" Foreground="Green" ToolTip="Active computers"/>
                                <TextBlock VerticalAlignment="Center" Text="/"/>
                                <TextBlock VerticalAlignment="Center" Name="txtPassiveComputerCount" Foreground="Red" ToolTip="Passive computers"/>
                                <TextBlock VerticalAlignment="Center" Text="]"/>
                            </StackPanel>
                        </TabItem.Header>
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <DataGrid Grid.Row="0" Name="computersDataGrid" CanUserAddRows="False" AutoGenerateColumns="False" SelectionMode="Single">
                                <DataGrid.RowHeaderTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal">
                                            <Path Height="10" Width="10" Margin="0,0,2,0" Stretch="Uniform">
                                                <Path.Style>
                                                    <Style TargetType="Path">
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource Mode=FindAncestor, AncestorType={x:Type DataGridRow}}, Path=Item.Enabled}" Value="True">
                                                                <Setter Property="Data" Value="{StaticResource enabledIcon}"/>
                                                                <Setter Property="Fill" Value="Green"/>
                                                                <Setter Property="ToolTip" Value="Enabled"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource Mode=FindAncestor, AncestorType={x:Type DataGridRow}}, Path=Item.Enabled}" Value="False">
                                                                <Setter Property="Data" Value="{StaticResource disabledIcon}"/>
                                                                <Setter Property="Fill" Value="Red"/>
                                                                <Setter Property="ToolTip" Value="Disabled"/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Path.Style>
                                            </Path>
                                        </StackPanel>
                                    </DataTemplate>
                                </DataGrid.RowHeaderTemplate>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem Name="ctxRDP" Header="Connect with Remote Desktop (RDP)"/>                                        
                                        <MenuItem Name="ctxMSRA" Header="Offer remote assistance (MSRA)"/>
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                            </DataGrid>
                            <Grid Grid.Row="1">
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Selected object CN:"/>
                                    <TextBlock Text="{Binding ElementName=computersDataGrid, Path=SelectedItem.Name}"/>
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                    <TextBlock Text="Selected object DN:"/>
                                    <TextBlock Text="{Binding ElementName=computersDataGrid, Path=SelectedItem.DistinguishedName}" TextWrapping="WrapWithOverFlow"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </TabItem>
                    <TabItem Name="tabItemUsers">
                        <TabItem.Header>
                            <StackPanel Orientation="Horizontal">
                                <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource userIcon}"/>
                                <TextBlock VerticalAlignment="Center" Text=" Users"/>
                                <TextBlock VerticalAlignment="Center" Text=" ["/>
                                <TextBlock VerticalAlignment="Center" Name="txtTotalUserCount" ToolTip="Active users"/>
                                <TextBlock VerticalAlignment="Center" Text="/"/>
                                <TextBlock VerticalAlignment="Center" Name="txtActiveUserCount" Foreground="Green" ToolTip="Active users"/>
                                <TextBlock VerticalAlignment="Center" Text="/"/>
                                <TextBlock VerticalAlignment="Center" Name="txtPassiveUserCount" Foreground="Red" ToolTip="Passive users"/>
                                <TextBlock VerticalAlignment="Center" Text="]"/>
                            </StackPanel>
                        </TabItem.Header>
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <DataGrid Grid.Row="0" Name="usersDataGrid" CanUserAddRows="False" AutoGenerateColumns="False" SelectionMode="Single">
                                <DataGrid.RowHeaderTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal">
                                            <Path Height="10" Width="10" Margin="0,0,2,0" Stretch="Uniform">
                                                <Path.Style>
                                                    <Style TargetType="Path">
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource Mode=FindAncestor, AncestorType={x:Type DataGridRow}}, Path=Item.Enabled}" Value="True">
                                                                <Setter Property="Data" Value="{StaticResource enabledIcon}"/>
                                                                <Setter Property="Fill" Value="Green"/>
                                                                <Setter Property="ToolTip" Value="Enabled"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding RelativeSource={RelativeSource Mode=FindAncestor, AncestorType={x:Type DataGridRow}}, Path=Item.Enabled}" Value="False">
                                                                <Setter Property="Data" Value="{StaticResource disabledIcon}"/>
                                                                <Setter Property="Fill" Value="Red"/>
                                                                <Setter Property="ToolTip" Value="Disabled"/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Path.Style>
                                            </Path>
                                        </StackPanel>
                                    </DataTemplate>
                                </DataGrid.RowHeaderTemplate>
                            </DataGrid>
                            <Grid Grid.Row="1">
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Selected object CN:"/>
                                    <TextBlock Text="{Binding ElementName=usersDataGrid, Path=SelectedItem.Name}"/>
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                    <TextBlock Text="Selected object DN:"/>
                                    <TextBlock Text="{Binding ElementName=usersDataGrid, Path=SelectedItem.DistinguishedName}" TextWrapping="WrapWithOverFlow"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </TabItem>
                </TabControl>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@
[xml]$script:xamlSettingsWindow = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window" WindowStartupLocation = "CenterScreen"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Width = "800" Height = "600" MinHeight="400" ShowInTaskbar = "True">
    <Window.Resources>        
        <Geometry x:Key="computerIcon">M6,2C4.89,2 4,2.89 4,4V12C4,13.11 4.89,14 6,14H18C19.11,14 20,13.11 20,12V4C20,2.89 19.11,2 18,2H6M6,4H18V12H6V4M4,15C2.89,15 2,15.89 2,17V20C2,21.11 2.89,22 4,22H20C21.11,22 22,21.11 22,20V17C22,15.89 21.11,15 20,15H4M8,17H20V20H8V17M9,17.75V19.25H13V17.75H9M15,17.75V19.25H19V17.75H15Z</Geometry>
        <Geometry x:Key="userIcon">M12,4A4,4 0 0,1 16,8A4,4 0 0,1 12,12A4,4 0 0,1 8,8A4,4 0 0,1 12,4M12,14C16.42,14 20,15.79 20,18V20H4V18C4,15.79 7.58,14 12,14Z</Geometry>
        <Geometry x:Key="minusIcon">M12,20C7.59,20 4,16.41 4,12C4,7.59 7.59,4 12,4C16.41,4 20,7.59 20,12C20,16.41 16.41,20 12,20M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M7,13H17V11H7</Geometry>
        <Geometry x:Key="plusIcon">M12,20C7.59,20 4,16.41 4,12C4,7.59 7.59,4 12,4C16.41,4 20,7.59 20,12C20,16.41 16.41,20 12,20M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M13,7H11V11H7V13H11V17H13V13H17V11H13V7Z</Geometry>
        <Geometry x:Key="settingsIcon">M12,15.5A3.5,3.5 0 0,1 8.5,12A3.5,3.5 0 0,1 12,8.5A3.5,3.5 0 0,1 15.5,12A3.5,3.5 0 0,1 12,15.5M19.43,12.97C19.47,12.65 19.5,12.33 19.5,12C19.5,11.67 19.47,11.34 19.43,11L21.54,9.37C21.73,9.22 21.78,8.95 21.66,8.73L19.66,5.27C19.54,5.05 19.27,4.96 19.05,5.05L16.56,6.05C16.04,5.66 15.5,5.32 14.87,5.07L14.5,2.42C14.46,2.18 14.25,2 14,2H10C9.75,2 9.54,2.18 9.5,2.42L9.13,5.07C8.5,5.32 7.96,5.66 7.44,6.05L4.95,5.05C4.73,4.96 4.46,5.05 4.34,5.27L2.34,8.73C2.21,8.95 2.27,9.22 2.46,9.37L4.57,11C4.53,11.34 4.5,11.67 4.5,12C4.5,12.33 4.53,12.65 4.57,12.97L2.46,14.63C2.27,14.78 2.21,15.05 2.34,15.27L4.34,18.73C4.46,18.95 4.73,19.03 4.95,18.95L7.44,17.94C7.96,18.34 8.5,18.68 9.13,18.93L9.5,21.58C9.54,21.82 9.75,22 10,22H14C14.25,22 14.46,21.82 14.5,21.58L14.87,18.93C15.5,18.67 16.04,18.34 16.56,17.94L19.05,18.95C19.27,19.03 19.54,18.95 19.66,18.73L21.66,15.27C21.78,15.05 21.73,14.78 21.54,14.63L19.43,12.97Z</Geometry>
        <Geometry x:Key="upIcon">M14,20H10V11L6.5,14.5L4.08,12.08L12,4.16L19.92,12.08L17.5,14.5L14,11V20Z</Geometry>
        <Geometry x:Key="downIcon">M10,4H14V13L17.5,9.5L19.92,11.92L12,19.84L4.08,11.92L6.5,9.5L10,13V4Z</Geometry>
        <SolidColorBrush x:Key="iconColor">#336699</SolidColorBrush>

        <sys:Double x:Key="FontSize">13</sys:Double>
        <sys:Double x:Key="ButtonSize">28</sys:Double>        
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TabControl Grid.Row="0" TabStripPlacement="Left" HorizontalContentAlignment="Left" Margin="5">
            <TabItem>
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                        <Grid MinWidth="25">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource settingsIcon}"/>
                        </Grid>
                        <TextBlock Text="General settings"/>
                    </StackPanel>
                </TabItem.Header>
                
                <Grid Margin="5" Name="grdGeneralSettings">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>                                                                                                
                    </Grid.RowDefinitions>
                    
                    <Grid Grid.Column="0" Grid.ColumnSpan="2" Grid.Row="0" Background="LightGray" Margin="0,2,0,4">
                        <TextBlock Text="General settings" Margin="2"/>
                    </Grid>
                    
                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Show Verbose output" Margin="0,2,0,0"/>
                    <CheckBox Grid.Row="1" Grid.Column="1" VerticalAlignment="Center" IsChecked="{Binding ShowVerboseOutput, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>
                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Show Debug output" Margin="0,2,0,0"/>
                    <CheckBox Grid.Row="2" Grid.Column="1" VerticalAlignment="Center" IsChecked="{Binding ShowDebugOutput, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>
                    <TextBlock Grid.Row="3" Grid.Column="0" Text="OnStart load OUs" Margin="0,2,0,0"/>
                    <CheckBox Grid.Row="3" Grid.Column="1" VerticalAlignment="Center"  IsChecked="{Binding OnStartLoadOrganizationalUnits, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>
                    <TextBlock Grid.Row="4" Grid.Column="0" Text="OnStart load groups" Margin="0,2,0,0"/>
                    <CheckBox Grid.Row="4" Grid.Column="1" VerticalAlignment="Center" IsChecked="{Binding OnStartLoadGroups, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>
                    <TextBlock Grid.Row="5" Grid.Column="0" Text="Computers inactive after [n] days  " Margin="0,2,0,0"/>
                    <TextBox Grid.Row="5" Grid.Column="1" Text="{Binding ComputerInactiveLimit, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>
                    <TextBlock Grid.Row="6" Grid.Column="0" Text="Users inactive after [n] days" Margin="0,2,0,0"/>
                    <TextBox Grid.Row="6" Grid.Column="1" Text="{Binding UserInactiveLimit, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Margin="0,2,0,0"/>                                                                                                                                   
                </Grid>
            </TabItem>
            
            <TabItem Name="tabItemcomputerAttributes">
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal">
                        <Grid MinWidth="25">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource computerIcon}"/>
                        </Grid>
                        <TextBlock Text="Computer attributes"/>
                    </StackPanel>
                </TabItem.Header>
                
                <Grid Margin="5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*" MinHeight="200"/>
                    </Grid.RowDefinitions>
                    
                    <Grid Grid.Column="0" Grid.ColumnSpan="2" Grid.Row="0" Background="LightGray" Margin="0,2,0,4">
                        <TextBlock Text="Define computer attribute definitions" Margin="2"/>
                    </Grid>
                    
                    <StackPanel Margin="0,0,2,0" Grid.Column="0" Grid.Row="1" Orientation="Vertical">
                        
                        <Button Background="Transparent" Name="btnAddComputerAttributeDefinition" ToolTip="Add attribute definition">
                            <Path Stretch="Uniform" Fill="Green"  Data="{StaticResource plusIcon}"/>
                        </Button>
                        
                        <Button Background="Transparent" Name="btnRemoveComputerAttributeDefinition" ToolTip="Remove attribute definition">
                            <Button.Style>
                                <Style TargetType="Button">
                                    <Setter Property="IsEnabled" Value="True"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding ElementName=dgComputerAttributes, Path=SelectedIndex}" Value="-1">
                                            <Setter Property="IsEnabled" Value="False"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </Button.Style>
                            <Path Stretch="Uniform" Fill="Red"  Data="{StaticResource minusIcon}"/>
                        </Button>
                        
                        <Button Margin="0,15,0,0" Name="btnUpComputerAttributeDefinition">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource upIcon}"/>
                        </Button>
                        
                        <Button Margin="0,5,0,0" Name="btnDownComputerAttributeDefinition">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource downIcon}"/>
                        </Button>
                        
                    </StackPanel>
                    <DataGrid Grid.Column="1" Grid.Row="1" ItemsSource="{Binding ComputerAttributeDefinitions}" Name="dgComputerAttributes" HeadersVisibility="Column" AutoGenerateColumns="False" CanUserAddRows="False">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Header="Attribute" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding Attribute, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"  VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            
                            <DataGridTemplateColumn Header="Friendly name" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding FriendlyName, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"  VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            
                            <DataGridTemplateColumn Header="Editable" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IsEditable, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                            <DataGridTemplateColumn Header="Ignore converter" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IgnoreConverter, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                            <DataGridTemplateColumn Header="Display in" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <ComboBox Text="{Binding DisplayIn, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                                            <ComboBoxItem Content="DataGrid"/>
                                            <ComboBoxItem Content="DetailsPane"/>
                                        </ComboBox>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
            
            <TabItem >
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal">
                        <Grid MinWidth="25">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource userIcon}"/>
                        </Grid>
                        <TextBlock Text="User attributes"/>
                    </StackPanel>
                </TabItem.Header>
                
                <Grid Margin="5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*" MinHeight="200"/>
                    </Grid.RowDefinitions>
                    
                    <Grid Grid.Column="0" Grid.ColumnSpan="2" Grid.Row="0" Background="LightGray" Margin="0,2,0,4">
                        <TextBlock Text="Define user attribute definitions" Margin="2"/>
                    </Grid>
                    
                    <StackPanel Margin="0,0,2,0" Grid.Column="0" Grid.Row="1" Orientation="Vertical">
                        
                        <Button Background="Transparent" Name="btnAddUserAttributeDefinition" ToolTip="Add attribute definition">
                            <Path Stretch="Uniform" Fill="Green"  Data="{StaticResource plusIcon}"/>
                        </Button>
                       
                        <Button Background="Transparent" Name="btnRemoveUserAttributeDefinition" ToolTip="Remove attribute definition">
                            <Button.Style>
                                <Style TargetType="Button">
                                    <Setter Property="IsEnabled" Value="True"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding ElementName=dgUserAttributes, Path=SelectedIndex}" Value="-1">
                                            <Setter Property="IsEnabled" Value="False"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </Button.Style>
                            <Path Stretch="Uniform" Fill="Red"  Data="{StaticResource minusIcon}"/>
                        </Button>

                        <Button Margin="0,15,0,0" Name="btnUpUserAttributeDefinition">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource upIcon}"/>
                        </Button>
                        
                        <Button Margin="0,5,0,0" Name="btnDownUserAttributeDefinition">
                            <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource downIcon}"/>
                        </Button>

                    </StackPanel>

                    <DataGrid Grid.Column="1" Grid.Row="1" ItemsSource="{Binding UserAttributeDefinitions}" Name="dgUserAttributes" HeadersVisibility="Column" AutoGenerateColumns="False" CanUserAddRows="False">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Header="Attribute" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding Attribute, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"  VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            
                            <DataGridTemplateColumn Header="Friendly name" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding FriendlyName, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"  VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            
                            <DataGridTemplateColumn Header="Editable" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IsEditable, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                             <DataGridTemplateColumn Header="Ignore converter" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IgnoreConverter, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                            <DataGridTemplateColumn Header="Display in" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <ComboBox Text="{Binding DisplayIn, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                                            <ComboBoxItem Content="DataGrid"/>
                                            <ComboBoxItem Content="DetailsPane"/>
                                        </ComboBox>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
        </TabControl>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Content="Ok" Name="btnOk" Margin="5" Height="{StaticResource ButtonSize}" MinWidth="60"/>
            <Button Content="Cancel" Name="btnCancel" Margin="5" Height="{StaticResource ButtonSize}" MinWidth="60"/>
        </StackPanel>

    </Grid>
</Window>
"@
[xml]$script:xamlExportWindow = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window" WindowStartupLocation = "CenterScreen"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    ShowInTaskbar = "True" SizeToContent="WidthAndHeight">
     <Window.Resources>
        <Geometry x:Key="selectAllIcon">M9,9H15V15H9M7,17H17V7H7M15,5H17V3H15M15,21H17V19H15M19,17H21V15H19M19,9H21V7H19M19,21A2,2 0 0,0 21,19H19M19,13H21V11H19M11,21H13V19H11M9,3H7V5H9M3,17H5V15H3M5,21V19H3A2,2 0 0,0 5,21M19,3V5H21A2,2 0 0,0 19,3M13,3H11V5H13M3,9H5V7H3M7,21H9V19H7M3,13H5V11H3M3,5H5V3A2,2 0 0,0 3,5Z</Geometry>
        <SolidColorBrush x:Key="iconColor">#336699</SolidColorBrush>

        <sys:Double x:Key="FontSize">13</sys:Double>
        <sys:Double x:Key="ButtonSize">28</sys:Double>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>

    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" MinWidth="100"/>
            <ColumnDefinition Width="Auto" MinWidth="100"/>
            <ColumnDefinition>
                <ColumnDefinition.Style>
                    <Style TargetType="ColumnDefinition">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding ElementName=rdBtnXLSX, Path=IsChecked}" Value="True">
                                <Setter Property="Width" Value="Auto"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding ElementName=rdBtnXLSX, Path=IsChecked}" Value="False">
                                <Setter Property="Width" Value="0"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </ColumnDefinition.Style>
            </ColumnDefinition>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Column="0" Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Columns to include" Margin="5,0,0,0"/>
            <Button Grid.Column="1" Name="btnSelectAll" HorizontalAlignment="Right" Margin="0,0,5,0" Background="Transparent" ToolTip="Select all">
                <Path Stretch="Uniform" Fill="{StaticResource iconColor}"  Data="{StaticResource selectAllIcon}"/>
            </Button>
        </Grid>

        <ListBox Grid.Column="0" Grid.Row="1" ItemsSource="{Binding}" MinHeight="150" MinWidth="150" SelectionMode="Multiple" Name="lstBoxColumnsToInclude" Margin="5">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding FriendlyName}" Tag="{Binding}"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>

        <TextBlock Grid.Column="1" Grid.Row="0" Text="Sort by" Margin="5,0,0,0"/>
        <ListBox Grid.Column="1" Grid.Row="1" Width="{Binding ElementName=lstBoxColumnsToInclude, Path=ActualWidth}" ItemsSource="{Binding ElementName=lstBoxColumnsToInclude, Path=SelectedItems}" MinHeight="150" MinWidth="150" SelectionMode="Single" Name="lstBoxSortBy" Margin="5">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding FriendlyName}" Tag="{Binding}"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>

        <TextBlock Grid.Column="2" Grid.Row="0" Text="Group by" Margin="5,0,0,0"/>
        <ListBox Grid.Column="2" Grid.Row="1" Width="{Binding ElementName=lstBoxColumnsToInclude, Path=ActualWidth}" ItemsSource="{Binding ElementName=lstBoxColumnsToInclude, Path=SelectedItems}" MinHeight="150" MinWidth="150" SelectionMode="Single" Name="lstBoxGroupBy" Margin="5">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding FriendlyName}" Tag="{Binding}"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>

        <StackPanel Grid.Column="0" Grid.Row="2" Grid.ColumnSpan="3" Orientation="Vertical" Margin="5,0,0,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Format: "/>
                <StackPanel Orientation="Vertical">

                    <RadioButton IsChecked="True" Content="XLSX" Name="rdBtnXLSX"/>
                    <RadioButton IsChecked="False" Name="rdBtnCSV">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="CSV"/>
                              <StackPanel Orientation="Horizontal">
                                    <StackPanel.Style>
                                    <Style TargetType="StackPanel">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding ElementName=rdBtnCSV, Path=IsChecked}" Value="False">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                    </StackPanel.Style>
                                <TextBlock Text="   Separator: "/>
                                <TextBox Text=";" Width="20" VerticalContentAlignment="Center" Name="txtBoxDelimiter"/>
                            </StackPanel>
                        </StackPanel>
                    </RadioButton>

                </StackPanel>
            </StackPanel>
        </StackPanel>

        <StackPanel Grid.Column="0" Grid.ColumnSpan="3" Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Content="Ok" Name="btnOk" Margin="5" Height="{StaticResource ButtonSize}" MinWidth="60">
                <Button.Style>
                    <Style TargetType="Button">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding ElementName=lstBoxSortBy, Path=SelectedIndex}" Value="-1">
                                <Setter Property="IsEnabled" Value="False"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </Button.Style>
            </Button>
            <Button Content="Cancel" Name="btnCancel" Margin="5" Height="{StaticResource ButtonSize}" MinWidth="60"/>
        </StackPanel>

    </Grid>
</Window>
"@
[xml]$script:xamllapsWindow = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window" WindowStartupLocation = "CenterScreen"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Width = "250" Height = "130" MinHeight="50" ShowInTaskbar = "True" ResizeMode="NoResize">
    <Window.Resources>            
        <Geometry x:Key="showPasswordIcon">F1 M 38,33.1538C 40.6765,33.1538 42.8462,35.3235 42.8462,38C 42.8462,40.6765 40.6765,42.8461 38,42.8461C 35.3235,42.8461 33.1539,40.6765 33.1539,38C 33.1539,35.3235 35.3236,33.1538 38,33.1538 Z M 38,25.0769C 49.3077,25.0769 59,33.1538 59,38C 59,42.8461 49.3077,50.9231 38,50.9231C 26.6923,50.9231 17,42.8461 17,38C 17,33.1538 26.6923,25.0769 38,25.0769 Z M 38,29.1154C 33.0932,29.1154 29.1154,33.0932 29.1154,38C 29.1154,42.9068 33.0932,46.8846 38,46.8846C 42.9068,46.8846 46.8846,42.9068 46.8846,38C 46.8846,33.0932 42.9068,29.1154 38,29.1154 Z </Geometry>    
        <SolidColorBrush x:Key="iconColor">#336699</SolidColorBrush>
        <sys:Double x:Key="FontSize">13</sys:Double>
        <sys:Double x:Key="ButtonSize">28</sys:Double>        
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
                        
        <TextBlock Grid.Column="0" Grid.Row="0" Text="Hostname: "/>
        <TextBox Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="0" Name="txtHostname" IsReadOnly="True"/>

        <TextBlock Grid.Column="0" Grid.Row="1" Text="Password: " Margin="0,2,0,0"/>
        <TextBox Grid.Column="1" Grid.Row="1" Name="txtPasswd" IsReadOnly="True" Margin="0,2,0,0"/>
        <Button Grid.Column="2" Grid.Row="1" Name="btnShowPassword" Height="20" Width="20" Margin="2,2,0,0">
            <Path Margin="2" Stretch="Uniform" Fill="{StaticResource iconColor}" Data="{StaticResource showPasswordIcon}"/>
        </Button>
        <Button Grid.Column="0" Grid.ColumnSpan="3" HorizontalAlignment="Right" Grid.Row="2" Content="Close" Width="60" Name="btnClose" Margin="0,5,0,0"/>                                    
    </Grid>
</Window>
"@
[xml]$script:xamldebugWindow = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window" WindowStartupLocation = "CenterScreen"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Width = "650" Height = "730" MinHeight="50" ShowInTaskbar = "True">
    <Window.Resources>            
        <Geometry x:Key="showPasswordIcon">F1 M 38,33.1538C 40.6765,33.1538 42.8462,35.3235 42.8462,38C 42.8462,40.6765 40.6765,42.8461 38,42.8461C 35.3235,42.8461 33.1539,40.6765 33.1539,38C 33.1539,35.3235 35.3236,33.1538 38,33.1538 Z M 38,25.0769C 49.3077,25.0769 59,33.1538 59,38C 59,42.8461 49.3077,50.9231 38,50.9231C 26.6923,50.9231 17,42.8461 17,38C 17,33.1538 26.6923,25.0769 38,25.0769 Z M 38,29.1154C 33.0932,29.1154 29.1154,33.0932 29.1154,38C 29.1154,42.9068 33.0932,46.8846 38,46.8846C 42.9068,46.8846 46.8846,42.9068 46.8846,38C 46.8846,33.0932 42.9068,29.1154 38,29.1154 Z </Geometry>    
        <SolidColorBrush x:Key="iconColor">#336699</SolidColorBrush>
        <sys:Double x:Key="FontSize">13</sys:Double>
        <sys:Double x:Key="ButtonSize">28</sys:Double>        
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="FontSize" Value="{StaticResource FontSize}" />
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="FontSize" Value="{StaticResource FontSize}"/>
        </Style>
    </Window.Resources>
    <DataGrid Name="dg" IsReadOnly="True"/>
</Window>
"@
#endregion

#region CSharp
$script:converters = @"
using System;
using System.Globalization;
using System.Windows.Data;
using Microsoft.ActiveDirectory.Management;
using System.Collections.Generic;

    public class ADPropertyValueCollectionConverter : IValueConverter
    {

        public string Separator
        {
            get { return ";"; }
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                string returnString = string.Empty;
                ADPropertyValueCollection collection = (ADPropertyValueCollection)value;
                if (collection.Count > 0)
                {
                    for(int i = 0; i < collection.Count; i++)
                    {
                        returnString = returnString + collection[i].ToString() + ";";
                    }
                }
                if(returnString.EndsWith(";"))
                {
                    returnString = returnString.Remove(returnString.Length - 1);
                }
                return returnString;
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class FileTimeConverter : IValueConverter
    {
        private string dateTimeFormat;

        public FileTimeConverter(string _dateTimeFormat)
        {
            this.dateTimeFormat = _dateTimeFormat;
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && (long)value > 0)
            {
                return DateTime.FromFileTime((long)value).ToString(dateTimeFormat);
            }
            else
            {
                return String.Empty;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class DateFormatConverter : IValueConverter
    {
        private string dateTimeFormat;

        public DateFormatConverter(string _dateTimeFormat)
        {
            this.dateTimeFormat = _dateTimeFormat;
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                DateTime dt = (DateTime)value;
                return dt.ToString(dateTimeFormat);
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class msExchRemoteRecipientTypeConverter : IValueConverter
    {
        Dictionary<Int64, string> dictionary = new Dictionary<Int64, string>()
        {
            {1,   "ProvisionMailbox"},
            {2,   "ProvisionArchive (On-Prem Mailbox)"},
            {3,   "ProvisionMailbox, ProvisionArchive"},
            {4,   "Migrated (UserMailbox)"},
            {6,   "ProvisionArchive, Migrated"},
            {8,   "DeprovisionMailbox"},
            {10,  "ProvisionArchive, DeprovisionMailbox"},
            {16,  "DeprovisionArchive (On-Prem Mailbox)"},
            {17,  "ProvisionMailbox, DeprovisionArchive"},
            {20,  "Migrated, DeprovisionArchive"},
            {24,  "DeprovisionMailbox, DeprovisionArchive"},
            {33,  "ProvisionMailbox, RoomMailbox"},
            {35,  "ProvisionMailbox, ProvisionArchive, RoomMailbox"},
            {36,  "Migrated, RoomMailbox"},
            {38,  "ProvisionArchive, Migrated, RoomMailbox"},
            {49,  "ProvisionMailbox, DeprovisionArchive, RoomMailbox"},
            {52,  "Migrated, DeprovisionArchive, RoomMailbox"},
            {65,  "ProvisionMailbox, EquipmentMailbox"},
            {67,  "ProvisionMailbox, ProvisionArchive, EquipmentMailbox"},
            {68,  "Migrated, EquipmentMailbox"},
            {70,  "ProvisionArchive, Migrated, EquipmentMailbox"},
            {81,  "ProvisionMailbox, DeprovisionArchive, EquipmentMailbox"},
            {84,  "Migrated, DeprovisionArchive, EquipmentMailbox"},
            {100, "Migrated, SharedMailbox"},
            {102, "ProvisionArchive, Migrated, SharedMailbox"},
            {116, "Migrated, DeprovisionArchive, SharedMailbox"}
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Int64)
            {
                if (dictionary.ContainsKey((Int64)value))
                {
                    return dictionary[(Int64)value];
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class msExchRecipientDisplayTypeConverter : IValueConverter
    {
        Dictionary<int, string> dictionary = new Dictionary<int, string>()
        {
            {-2147483642,   "MailUser (RemoteUserMailbox)"},
            {-2147481850,   "MailUser (RemoteRoomMailbox)"},
            {-2147481594,   "MailUser (RemoteEquipmentMailbox)"},
            {0,             "UserMailbox (shared)"},
            {1,             "MailUniversalDistributionGroup"},
            {6,             "MailContact"},
            {7,             "UserMailbox (room)"},
            {8,             "UserMailbox (equipment)"},
            {1073741824 ,   "UserMailbox"},
            {1073741833 ,   "MailUniversalSecurityGroup"}
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int)
            {
                if (dictionary.ContainsKey((int)value))
                {
                    return dictionary[(int)value];
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class managedByConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string)
            {
                string val = (string)value;
                if (val.Length > 4 && val.Contains(","))
                {
                    string retValue = (string)val;
                    retValue = retValue.Substring(3);
                    retValue = retValue.Substring(0, retValue.IndexOf(','));
                    return retValue;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
"@
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
#######################################################################
#                                                                     #
# Merged by user: administrator                                       #
# On computer:    DC01                                                #
# Date:           2021-02-13 23:23:42                                 #
# No code signing certificate found!                                  #
# LoC: 3538 of which 2933 is not comments or whitespace               #
#                                                                     #
#######################################################################

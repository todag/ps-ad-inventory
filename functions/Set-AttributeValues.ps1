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
        Write-Error "Cannot commit, unknown target type"
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
        if((!$existingValue -and $newValue) -or ($existingValue -and !$newValue) -or ($existingValue -and $newValue -and (Compare-Object $newValue $existingValue)))
        {
            if($newValue -eq $null)
            {
                Write-Verbose($verboseStr + " pending action [clear]")
                $commitCount++
            }
            else
            {
                Write-Verbose($verboseStr + " pending action [new value]")
                $commitCount++
            }
            $TargetObject.($attr.Attribute) = $newValue
            $commitChanges = $true
        }
        else
        {
            Write-Verbose($verboseStr + " pending action [no changes]")
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
            Write-Verbose("[" + $TargetObject.Name + "] " + $commitCount.ToString() + " change(s) commited...")
        }
        catch
        {
            Write-Error $_.Exception.Message
            [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error') | Out-Null
        }
    }
    else
    {
        Write-Verbose("[" + $TargetObject.Name + "] " + $commitCount.ToString() + " change(s) commited...")
    }

    Write-Verbose("---------------------------------------------------------------------------------------")
}
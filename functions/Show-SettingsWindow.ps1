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
        Write-Verbose $settings.ComputerAttributeDefinitions[0].FriendlyName
        Write-Verbose $settingsUnconfirmed.ComputerAttributeDefinitions[0].FriendlyName

        if(Resolve-AttributeDefinitionsFromSchema -ComputerAttributeDefinitions $settingsUnconfirmed.ComputerAttributeDefinitions -UserAttributeDefinitions $settingsUnconfirmed.UserAttributeDefinitions)
        {
            $script:settings = $settingsUnconfirmed
            Export-Clixml -Path .\settings.xml -InputObject $script:settings
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
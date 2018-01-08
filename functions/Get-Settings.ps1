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
        $settingsImport = Import-Clixml -Path .\settings.xml
        Write-Verbose "Settings loaded from file. Generating Settings object"
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
        Write-ADIDebug ("Returning Settings object from file")
        return $settings
    }
    catch
    {
        #
        # Ok, something went wrong loading settings file. Generate new settings and as an extra
        # precaution, pass the definitions through Resolve-AttributeDefintionsFromSchema
        #
        Write-Verbose $_.Exception.Message
        Write-Verbose "Unable to load settings from file, using defaults"
        $newSettings = New-Object Settings
        $newSettings.ComputerAttributeDefinitions = Get-DefaultAttributeDefinitions -Computer
        $newSettings.UserAttributeDefinitions = Get-DefaultAttributeDefinitions -User

        if(Resolve-AttributeDefinitionsFromSchema -ComputerAttributeDefinitions $newSettings.ComputerAttributeDefinitions -UserAttributeDefinitions $newSettings.UserAttributeDefinitions)
        {
            Write-ADIDebug "Failed to get settings object from file, returning default settings object"
            return $newSettings
        }
        else
        {
            Write-Error "Unable to get default settings... script will terminate. Press enter to exit."
            [System.Windows.MessageBox]::Show("Unable to get default settings... script will terminate." + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
            Read-Host
            exit

        }
    }
}
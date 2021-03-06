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
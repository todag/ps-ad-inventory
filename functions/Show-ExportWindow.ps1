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
    foreach($guiObject in $xamlExportWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        Write-ADIDebug("Adding " + $guiObject.Name + " to $exportWindow")
        $exportWindow.$($guiObject.Name) = $exportWindow.Window.FindName($guiObject.Name)
    }

    $exportWindow.Window.DataContext = $AttributeDefinitions

    function Export-Inventory
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        $SaveFileDialog.Filter = “Text files (*.csv)|*.csv|All files (*.*)|*.*”
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
                }
                $exportObjects.Add($exportObj) | Out-Null
            }
            $delimiter = $exportWindow.txtBoxDelimiter.Text[0]
            if([string]::IsNullOrWhiteSpace($delimiter))
            {
                Write-ADIDebug "User defined delimiter is null or empty. Setting delimiter to ';'"
                $delimiter = ";"
            }
            else
            {
                Write-ADIDebug ("Delimiter is '" + $delimiter + "'")
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

            $exportObjects | Select-Object ($AttributesToExport).FriendlyName | Sort-Object -Property $exportWindow.lstBoxSortBy.SelectedItem.FriendlyName | Export-CSV -Path $SaveFileDialog.filename -Delimiter $delimiter -NoTypeInformation -Encoding UTF8
            Invoke-Item $SaveFileDialog.filename
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
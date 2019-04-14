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
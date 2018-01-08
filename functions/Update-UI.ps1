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

            Write-ADIDebug("Adding binding: "  + $attr.Attribute)
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
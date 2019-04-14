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



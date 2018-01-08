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
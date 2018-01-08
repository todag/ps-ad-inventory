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
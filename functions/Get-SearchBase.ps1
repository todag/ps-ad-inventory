###############################################################################
#.SYNOPSIS
#    Returns searchbase from either OU Browser or Groups Listbox (depending on which has a selected item)
#
###############################################################################
function Get-SearchBase
{
    if($script:MainWindow.tvOUBrowser.SelectedItem)
    {
        return $script:MainWindow.tvOUBrowser.SelectedItem.Tag
    }
    elseif($script:MainWindow.lstBoxGroups.SelectedIndex -gt -1)
    {
        return $script:MainWindow.lstBoxGroups.SelectedItem
    }
}
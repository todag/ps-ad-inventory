###############################################################################
#.SYNOPSIS
#    Returns searchscope from either OU Browser or Groups Listbox (depending on which has a selected item)
#
###############################################################################
function Get-SearchScope
{
    if((Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or (Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
    {
        if($script:MainWindow.chkRecursiveOUSearch.IsChecked)
        {
            return "Subtree"
        }
        else
        {
            return "OneLevel"
        }
    }
    elseif((Get-SearchBase).GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
    {
        if($script:MainWindow.chkRecursiveGroupSearch.IsChecked)
        {
            return "Subtree"
        }
        else
        {
            return "OneLevel"
        }
    }
}

###############################################################################
#.SYNOPSIS
#    Returns a TreeViewItem with children containing the domains OU structure.
#    The ADOrganizationalUnit object is stored in the items tag
#    The root items tag will contain a ADDomain object
#
#.PARAMETER SearchString
#    Will return results based on SearchString
#
###############################################################################
function Get-DomainTree
{
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory=$false)]
        [string]$SearchString
    )

    function Get-OUTreeChildItems()
    {
        [CmdletBinding()]
        Param
        (
            [parameter(Mandatory=$true)]
            [System.Windows.Controls.TreeViewItem] $ParentItem
        )

        foreach($ou in Get-ADOrganizationalUnit -Filter * -SearchBase $ParentItem.Tag -SearchScope OneLevel)
        {
            $script:ouCount++
            $treeViewItem = New-Object System.Windows.Controls.TreeViewItem
            $treeViewItem.Tag = $ou
            $treeViewItem.Header = $ou.Name
            Get-OUTreeChildItems -ParentItem $treeViewItem
            $ParentItem.Items.Add($treeViewItem) | Out-Null
        }
    }

    Write-Verbose "Loading Organizational units..."
    $startTime = (Get-Date)
    try
    {
        $rootItem = New-Object System.Windows.Controls.TreeViewItem
        $domain = Get-ADDomain
        $rootItem.Header = $domain.NetBIOSName
        $rootItem.Tag = $domain
        $script:ouCount = 0
        if([string]::IsNullOrWhiteSpace($SearchString))
        {
            Get-OUTreeChildItems -ParentItem $rootItem
        }
        else
        {
            $SearchString = ("*" + $SearchString + "*")
            Write-Verbose ("Loading Organizational units matching '" + $SearchString + "'")
            foreach($ou in Get-ADOrganizationalUnit -Filter { Name -like $SearchString } )
            {
                $script:ouCount++
                Write-ADIDebug ("Found " + $ou.Name)
                $treeViewItem = New-Object System.Windows.Controls.TreeViewItem
                $treeViewItem.Tag = $ou
                $treeViewItem.Header = $ou.Name
                $rootItem.Items.Add($treeViewItem) | Out-Null
            }
            $rootItem.Header = ($rootItem.Header + " (" + $script:ouCount.ToString() + ") search results")
        }
        Write-Verbose ($script:ouCount.ToString() + " units found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds")
    }
    catch
    {
        Write-Error $_.Exception.Message
    }
    $rootItem.IsExpanded = $true
    return $rootItem
}

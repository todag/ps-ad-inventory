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

    Write-Log -LogString "Loading Organizational units..." -Severity "Informational"
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
            Write-Log -LogString ("Loading Organizational units matching '" + $SearchString + "'") -Severity "Informational"
            foreach($ou in Get-ADOrganizationalUnit -Filter { Name -like $SearchString } )
            {
                $script:ouCount++
                Write-Log -LogString ("Found " + $ou.Name) -Severity "Debug"
                $treeViewItem = New-Object System.Windows.Controls.TreeViewItem
                $treeViewItem.Tag = $ou
                $treeViewItem.Header = $ou.Name
                $rootItem.Items.Add($treeViewItem) | Out-Null
            }
            $rootItem.Header = ($rootItem.Header + " (" + $script:ouCount.ToString() + ") search results")
        }
        Write-Log -LogString ($script:ouCount.ToString() + " units found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Critical"
    }
    $rootItem.IsExpanded = $true
    return $rootItem
}

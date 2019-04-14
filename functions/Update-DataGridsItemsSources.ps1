###############################################################################
#.SYNOPSIS
#    Loads ADComputer and ADUser objects from Active Directory and
#    sets resulting collection as ItemsSource for the DataGrids
#
#.PARAMETER SearchBase
#    SearchBase can be either ADOrganizationalUnit, ADGroup or ADDomain
#
#.PARAMETER SearchString
#    Will filter results based on SearchString
#
#.PARAMETER SearchScope
#    SubTree or OneLevel
#
###############################################################################
function Update-DataGridsItemsSources
{
    Param
    (
        [Parameter(Mandatory=$false)]
        [Object]$SearchBase = (Get-SearchBase),
        [Parameter(Mandatory=$false)]
        [string]$SearchString = $null,
        [Parameter(Mandatory=$false)]
        [string]$SearchScope = (Get-SearchScope)
    )

    #
    # Generate -LDAPFilter string. This is more flexible then using regular -Filter
    #
    if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
    {
        $computerLdapFilter = "(&(objectCategory=computer)"
        $userLdapFilter = "(&(objectCategory=person)"
    }
    elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
    {
        if($SearchScope -eq "Subtree")
        {
            $computerLdapFilter = "(&(objectCategory=computer)(memberof:1.2.840.113556.1.4.1941:=" + $SearchBase.DistinguishedName + ")"
            $userLdapFilter = "(&(objectCategory=person)(memberof:1.2.840.113556.1.4.1941:=" + $SearchBase.DistinguishedName + ")"
        }
        else
        {
            $computerLdapFilter = "(&(objectCategory=computer)(memberof=" + $SearchBase.DistinguishedName + ")"
            $userLdapFilter = "(&(objectCategory=person)(memberof=" + $SearchBase.DistinguishedName + ")"
        }
    }
    else
    {
        Write-Log -LogString "Error, unknown type as SearchBase!" -Severity "Error"
        return
    }

    #
    # If user has provided a filter string, continue generation of the -LDAPFilter string
    #
    if(![string]::IsNullOrWhiteSpace($SearchString))
    {
        $computerLdapFilter = $computerLdapFilter + "(|(cn=*" + $SearchString + "*)"
        $userLdapFilter = $userLdapFilter + "(|(cn=*" + $SearchString + "*)"
        foreach($attr in $script:settings.ComputerAttributeDefinitions)
        {
            $computerLdapFilter = ($computerLdapFilter + "(" + $attr.Attribute + "=*" + $SearchString + "*)")
        }
        foreach($attr in $script:settings.UserAttributeDefinitions)
        {
            $userLdapFilter = ($userLdapFilter + "(" + $attr.Attribute + "=*" + $SearchString + "*)")
        }
        $computerLdapFilter = $computerLdapFilter + ")"
        $userLdapFilter = $userLdapFilter + ")"
    }

    $computerLdapFilter = $computerLdapFilter + ")"
    $userLdapFilter = $userLdapFilter + ")"


    #
    # Fetch computer objects.
    #
    $startTime = Get-Date
    try
    {
        Write-Log -LogString ("Loading computers from " + $SearchBase.GetType()) -Severity "Informational"
        if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
        {
            $script:MainWindow.computersDataGrid.ItemsSource = @(Get-ADComputer -LDAPFilter $computerLdapFilter -SearchScope $SearchScope -SearchBase $SearchBase.DistinguishedName -Properties $script:settings.ComputerAttributeDefinitions.Attribute)
        }
        elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
        {
            $script:MainWindow.computersDataGrid.ItemsSource = @(Get-ADComputer -LDAPFilter $computerLdapFilter -Properties $script:settings.ComputerAttributeDefinitions.Attribute)
        }
        Write-Log -LogString ($script:MainWindow.computersDataGrid.ItemsSource.Count.ToString() + " found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error')
    }

    #
    # Fetch user objects.
    #
    $startTime = Get-Date
    try
    {
        Write-Log -LogString ("Loading users from " + $SearchBase.GetType()) -Severity "Informational"
        [System.Collections.ObjectModel.ObservableCollection[Object]]$usersCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        if($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADOrganizationalUnit] -or $SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADDomain])
        {
            $script:MainWindow.usersDataGrid.ItemsSource = @(Get-ADUser -LDAPFilter $userLdapFilter -SearchScope $SearchScope -SearchBase $SearchBase.DistinguishedName -Properties $script:settings.UserAttributeDefinitions.Attribute)
        }
        elseif($SearchBase.GetType() -eq [Microsoft.ActiveDirectory.Management.ADGroup])
        {
            $script:MainWindow.usersDataGrid.ItemsSource = @(Get-ADUser -LDAPFilter $userLdapFilter -Properties $script:settings.UserAttributeDefinitions.Attribute)
        }
        Write-Log -LogString ($script:MainWindow.usersDataGrid.ItemsSource.Count.ToString() + " found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Exception",'Ok','Error')
    }

    #
    # Select the TabItem containing the highest items count
    #
    if($script:MainWindow.computersDataGrid.ItemsSource.Count -gt $script:MainWindow.usersDataGrid.ItemsSource.Count)
    {
        $script:MainWindow.tabItemComputers.IsSelected = $true
    }
    elseif($script:MainWindow.usersDataGrid.ItemsSource.Count -gt $script:MainWindow.computersDataGrid.ItemsSource.Count)
    {
        $script:MainWindow.tabItemUsers.IsSelected = $true
    }

    #Update some UI controls...
    Update-UI
    Update-Statistics
}
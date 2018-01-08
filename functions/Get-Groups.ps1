###############################################################################
#.SYNOPSIS
#    Returns Groups from Active Directory
#
#.PARAMETER SearchString
#    Filters which groups to return based on SearchString
#
###############################################################################
function Get-Groups()
{
    Param
    (
        [parameter(Mandatory=$false)]
        [string]$SearchString
    )
    Write-Verbose "Loading Groups..."
    $startTime = (Get-Date)
    try
    {
        if(!$SearchString)
        {
            $result = @(Get-ADGroup -Filter * | Sort-Object Name)
        }
        else
        {
            $filter = "*" + $SearchString + "*"
            $result = @(Get-ADGroup -Filter { Name -like $filter } | Sort-Object Name)
        }

        Write-Verbose ($result.Count.ToString() + " groups found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds")
        return ,$result
    }
    catch
    {
        Write-Error $_.Exception.Message
    }

}

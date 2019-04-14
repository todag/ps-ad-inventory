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
    Write-Log -LogString "Loading Groups..." -Severity "Informational"
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

        Write-Log -LogString ($result.Count.ToString() + " groups found in " + ((Get-Date) - $startTime).TotalSeconds + " seconds") -Severity "Informational"
        return ,$result
    }
    catch
    {
        Write-Log -LogString $_.Exception.Message -Severity "Critical"
    }

}

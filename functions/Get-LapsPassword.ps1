###############################################################################
#
#.SYNOPSIS
#    Returns the LAPS password from Active Directory
#
#.PARAMETER Computer
#    Will return the LAPS password for this computer
#
###############################################################################
function Get-LapsPassword
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADComputer]$Computer        
    )
    Write-Log -LogString ("Trying to get LAPS password for computer [" + $Computer.Name + "]") -Severity "Notice"
    
    $lapspwd = Get-ADComputer -Identity $Computer.DistinguishedName -Properties ms-mcs-admpwd | Select-Object -Expand ms-mcs-admpwd        
    if(![string]::IsNullOrEmpty($lapspwd))
    {
        return $lapspwd
    }
    else
    {        
        try
        {
            $lapspwd = Get-ADComputer -Credential (Get-Credential -Message "Password looks empty, you probably don't have permissions to the confidential attribute. Specify different account with sufficient permissions to try again") -Identity $Computer.DistinguishedName -Properties ms-mcs-admpwd | Select-Object -Expand ms-mcs-admpwd            
            if(![string]::IsNullOrEmpty($lapspwd))
            {
                return $lapspwd
            }
            else
            {
                [System.Windows.MessageBox]::Show("The retrieved password is empty, you might have insufficient permissions to read the attribute: ", "Error",'Ok','Error') | Out-Null
            }
            
        }
        catch
        {
            Write-Log -LogString ("Exception retrieving LAPS password: " + $_.Exception.Message) -Severity "Error"
            [System.Windows.MessageBox]::Show("Exception retrieving LAPS password: " + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
            return $null
        }                     
    }    
}
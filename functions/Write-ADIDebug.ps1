###############################################################################
#.SYNOPSIS
#    Custom function to get function name prepended to debug output
#
#.PARAMETER DebugString
#    String to be written as debug output
#
#.EXAMPLE
#    Write-ADIDebug "DebugString"
#
###############################################################################
function Write-ADIDebug
{
    Param(
        [Parameter(Mandatory=$false)]
            $DebugString
        )

        $DebugString = "[" + $((Get-PSCallStack)[1].Command) + "]: " + $DebugString
        Write-Debug $DebugString
}
###############################################################################
#.SYNOPSIS
#    Checks AD schema and sanitizes user provided AttributeDefinitions
#
#.DESCRIPTION
#    Checks AD schema and sanitizes user provided AttributeDefinitions
#    If user has set an attribute as editable with Syntax other than DirectoryString
#    We will override and set it as not editable.
#
#.PARAMETER ComputerAttributeDefinitions
#    ComputerAttributeDefinitions to check
#
#.PARAMETER UserAttributeDefinitions
#    UserAttributeDefinitions to check
#
#.NOTES
#    Some special handling here for 'description' attribute. It's marked as multivalued in schema, when it's mostly not...
#    From MSDN:
#    Remarks
#    The description attribute is implemented as a multi-valued attribute in the schema for the cases where that is allowed. For an object that is not a SAM
#    managed class, the description is multi-valued. For an attribute that is a SAM managed class, the description attribute is single-valued. SAM managed
#    classes are for things like security principals so, if you have, for example, a container, or a class of your own, the schema will let you use multiple values.
#    This behavior of the description attribute is for backward compatibility with earlier operating systems because the attribute existed in the SAM APIs
#    before AD existed.
#    https://msdn.microsoft.com/en-us/library/ms675492%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
#
###############################################################################
function Resolve-AttributeDefinitionsFromSchema
{
    Param
    (
        [parameter(Mandatory=$true)]
        [System.Collections.Generic.List[AttributeDefinition]]$ComputerAttributeDefinitions,
        [parameter(Mandatory=$true)]
        [System.Collections.Generic.List[AttributeDefinition]]$UserAttributeDefinitions
    )

    function Resolve()
    {
        Param
        (
            [parameter(Mandatory=$true)]
            [System.Collections.Generic.List[AttributeDefinition]]$AttributeDefinitions,
            [parameter(Mandatory=$true)]
            [string]$ObjectCategory
        )

        foreach($attrDef in $AttributeDefinitions)
        {
            try
            {
                $schemaProperty = $schema.FindProperty($attrDef.Attribute)
                #Check if Single or MultiValue
                if($attrDef.Attribute -eq "description")
                {
                    Write-ADIDebug("[" + $ObjectCategory + "] '" + $attrDef.Attribute + "' setting as SingleValued <--- *** DEFINITION CHANGED ***")
                    $attrDef.IsSingleValued = $true
                }
                else
                {
                    $attrDef.IsSingleValued = $schemaProperty.IsSingleValued

                }
                #Get syntax
                $attrDef.Syntax = $schemaProperty.Syntax

                #Finally set AttributeDefinition to read-only if Syntax is not DirectoryString
                if($attrDef.Syntax -ne "DirectoryString" -and $attrDef.IsEditable -eq $true)
                {
                    $attrDef.IsEditable = $false
                    Write-ADIDebug("[" + $ObjectCategory + "] '" + $attrDef.Attribute +  "' Syntax:" + $attrDef.Syntax + " setting as Read-Only <--- *** DEFINITION CHANGED ***")
                }
            }
            catch
            {
                Write-Error $_.Exception.Message
                [System.Windows.MessageBox]::Show("Cannot resolve attribute '" + $attrDef.Attribute + "' from schema. " + $_.Exception.Message, "Error",'Ok','Error') | Out-Null
                return $false
            }

            Write-ADIDebug("[" + $ObjectCategory + "] '" + $attrDef.Attribute +  "' Syntax:" + $attrDef.Syntax + " IsSingleValued:" + $attrDef.IsSingleValued)
        }
        return $true
    }

    $schema = [DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
    $startTime = Get-Date
    Write-Verbose "Resolving AttributeDefinitions from schema..."
    $computerResolveResult = Resolve -AttributeDefinition $ComputerAttributeDefinitions -ObjectCategory "computer"
    $userResolveResult = Resolve -AttributeDefinition $UserAttributeDefinitions -ObjectCategory "user"

    Write-ADIDebug ("Computer attributes resolve result:: " + $computerResolveResult.ToString())
    Write-ADIDebug ("User attributes resolve result: " + $userResolveResult.ToString())

    Write-Verbose ("Resolved in " + ((Get-Date) - $startTime).TotalSeconds + " seconds")

    if($computerResolveResult -eq $true -and $userResolveResult -eq $true)
    {
        return $true
    }
    else
    {
        return $false
    }

}

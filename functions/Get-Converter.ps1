###############################################################################
#.SYNOPSIS
#    Returns an IValueConverter for an attribute.
#.DESCRIPTION
#    Currently contains converters for:
#    DateFormatConverter (for attributes with Syntax GeneralizedTime)
#    FileTimeConverter (for attributes that stores Date as FileTime)
#    ADPropertyValueCollectionConverter (For Multivalue attributes)
#    msExchRemoteRecipientTypeConverter (For attribute msExchRemoteRecipientTypeConverter)
#
#.PARAMETER AttributeDefinition
#    AttributeDefinition of the attribute to get the converter for
#
################################################################################
function Get-Converter
{
    Param
    (
        [parameter(Mandatory=$true)]
        [AttributeDefinition]$AttributeDefinition
    )

    $fileTimeAttributes = (
        "badPasswordTime",
        "lastLogonTimeStamp",
        "pwdLastSet",
        "accountExpires",
        "lastLogon",
        "lastPwdSet",
        "lockoutTime"
        )

    function Get-ConverterInstance
    {
        Param
        (
            [parameter(Mandatory=$true)]
            [string]$Converter,
            [parameter(Mandatory=$false)]
            $Param1
        )
        if((Get-Variable -Scope Script $Converter).Value -eq $null)
        {
            Write-ADIDebug ("Instantiating and returning IValueConverter: [" + $Converter + "][" + $AttributeDefinition.Attribute + "]")
            (Get-Variable -Scope Script $Converter).Value = New-Object -TypeName $Converter -ArgumentList $Param1
            return (Get-Variable -Scope Script $Converter).Value
        }
        else
        {
            Write-ADIDebug ("Returning IValueConverter: [" + $Converter + "][" + $AttributeDefinition.Attribute + "]")
            return (Get-Variable -Scope Script $Converter).Value
        }
    }

    if($AttributeDefinition.IgnoreConverter)
    {
        Write-ADIDebug("Ignoring converter for attribute '" + $AttributeDefinition.Attribute + "'")
        return $null
    }

    if($fileTimeAttributes -Contains ($AttributeDefinition.Attribute))
    {
        return Get-ConverterInstance -Converter "FileTimeConverter" -Param1 $script:dateTimeFormat
    }
    elseif($AttributeDefinition.Attribute -eq "managedBy")
    {
        return Get-ConverterInstance -Converter "managedByConverter"
    }
    elseif($AttributeDefinition.IsSingleValued -and $AttributeDefinition.Syntax -eq "GeneralizedTime")
    {
        return Get-ConverterInstance -Converter "DateFormatConverter" -Param1 $script:dateTimeFormat
    }
    elseif(!$AttributeDefinition.IsSingleValued)
    {
        return Get-ConverterInstance -Converter "ADPropertyValueCollectionConverter"
    }
    elseif($AttributeDefinition.Attribute -eq "msExchRemoteRecipientType")
    {
        return Get-ConverterInstance -Converter "msExchRemoteRecipientTypeConverter"
    }
    elseif($AttributeDefinition.Attribute -eq "msExchRecipientDisplayType")
    {
        return Get-ConverterInstance -Converter "msExchRecipientDisplayTypeConverter"
    }
}
###############################################################################
#
#.SYNOPSIS
#    Returns default attribute definitions for computer or user objects
#    These are only used on first run or if settings.xml is missing
#
#.PARAMETER Computer
#    Returns default attribute definitions for computer objects
#
#.PARAMETER User
#    Returns default attribute definitions for user objects
#
###############################################################################
function Get-DefaultAttributeDefinitions
{
    Param
    (
        [Parameter(Mandatory=$false)]
        [switch]$Computer,
        [Parameter(Mandatory=$false)]
        [switch]$User
    )

    function Get-DefaultComputerAttributeDefinitions
    {
        $computerAttributes = New-Object System.Collections.Generic.List[AttributeDefinition]
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Name",                "name",                     $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Company",             "company",                  $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Division",            "division",                 $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department",          "department",               $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department Number",   "departmentNumber",         $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Location",            "location",                 $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Description",         "description",              $true,   "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last user",           "info",                     $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Computer model",      "employeeType",             $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Operating system",    "operatingSystem",          $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "OS version",          "operatingSystemVersion",   $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last logged on",      "lastLogonTimestamp",       $false,  "DataGrid"))
        $computerAttributes.Add((New-Object AttributeDefinition -ArgumentList "Created",             "whenCreated",              $false,  "DataGrid"))
        return $computerAttributes
    }

    function Get-DefaultUserAttributeDefinitions
    {
        $userAttributes = New-Object System.Collections.Generic.List[AttributeDefinition]
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Name",                "name",               $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Company",             "company",            $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Department",          "department",         $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Title",               "title",              $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Description",         "description",        $true,   "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last logged on",      "lastLogonTimestamp", $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "UPN",                 "userPrincipalName",  $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Created",             "whenCreated",        $false,  "DataGrid"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Password last set",   "pwdLastSet",         $false,  "DetailsPane"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Last bad password",   "badPasswordTime",    $false,  "DetailsPane"))
        $userAttributes.Add((New-Object AttributeDefinition -ArgumentList "Bad password count",  "badPwdCount",        $false,  "DetailsPane"))
        return $userAttributes
    }

    if($Computer)
    {
        return Get-DefaultComputerAttributeDefinitions
    }
    elseif($User)
    {
        return Get-DefaultUserAttributeDefinitions
    }

}
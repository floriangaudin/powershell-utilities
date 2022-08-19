<#
    .SYNOPSIS
    Performs audit on inactive Active Directory's objects 

    .DESCRIPTION
    The inactive-activedirectory-object.ps1 script show an output or generate a csv-based file.

    .PARAMETER ADObject
    Specifies the type of Active Directory object based search

    .PARAMETER From
    Specifies the number of days for which the Active Directory 
    object has been disabled

    .PARAMETER ExportTo
    Specifies the path for the CSV-based output file. By default,
    inactive-activedirectory-object.ps1 generates an output.

    .INPUTS
    None. You cannot pipe objects to inactive-activedirectory-object.ps1.

    .OUTPUTS
    inactive-activedirectory-object.ps1 can generate output.

    .EXAMPLE
    PS> .\inactive-activedirectory-object.ps1 -ADObject Computer -From 180 -ExportTo C:\audits\

    .EXAMPLE
    PS> .\inactive-activedirectory-object.ps1 -ADObject User -From 365 -ExportTo C:\audits\
#>
Param(
    [parameter(Mandatory)]
    [ValidateSet('Computer','User')]
    [string]$ADObject,

    [parameter(Mandatory)]
    [ValidateScript({
        if( $_ -le 0) {
            throw "Must be higher of 0"
        }
        return $true
    })]
    [int]$From,

    [parameter()]
    [ValidateScript({
        if( -Not ($_ | Test-Path) ){
            throw "Folder does not exist"
        }
        return $true
    })]
    [System.IO.FileInfo]$ExportTo = $null
)

#remove number of days passed in arg to the current date
$calculatedDate = (Get-Date).Adddays(-($From))


switch ($ADObject) {

    Computer {  
        $results = Get-ADComputer -Filter {LastLogonTimeStamp -lt $calculatedDate} -ResultPageSize 2000 -resultSetSize $null -Properties Name, OperatingSystem, SamAccountName, DistinguishedName, description
    }
    user {
        $results = Get-ADUser -Filter {LastLogonTimeStamp -lt $calculatedDate} -ResultPageSize 2000 -resultSetSize $null -Properties Name, SamAccountName, DistinguishedName, description 
    }

}

if($ExportTo -eq $null){

    #-ExportTop is not specified, generate result's output on host
    $results | Format-Table -AutoSize

}else{

    #-ExportTop is specified, generate result's output in csv-based file
    $results | Export-CSV -Path ('{0}\inactive-{1}s-from-{2}.csv' -f $ExportTo, $ADObject, $calculatedDate.ToString("dd-MM-yyyy")) –NoTypeInformation

}

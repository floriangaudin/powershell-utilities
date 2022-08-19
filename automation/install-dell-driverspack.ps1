<#
    .SYNOPSIS
    Performs installation of Dell drivers pack

    .DESCRIPTION
    The install-dell-driverspack.ps1 install all drivers packed in .cab file

    .INPUTS
    None.

    .OUTPUTS
    None.

    .EXAMPLE
    PS> .\install-dell-driverspack.ps1
#>

Param(
    [parameter()]
    [ValidateScript({
        if( -Not ($_ | Test-Path) ){
            throw "Folder or file does not exist"
        }
        if($_ -notmatch "(\.cab)"){
            throw "The file specified in the path argument must be cab type"
        }
        return $true
    })]
    [System.IO.FileInfo]$File = $null
)

$tempDirectory = "C:\Temp\Cab"
$originalWorkingDirectory = $PSScriptRoot

if($File -eq $null){
    $workingDirectory = $originalWorkingDirectory
}else{
    $workingDirectory = [System.IO.Path]::GetDirectoryName($File)
}

Set-Location $workingDirectory

#if many, find last cab file added by CreationTime attribute
$cabFile = Get-ChildItem -Path .\ -Filter *.cab | Sort-Object  Name, CreationTime, LastWriteTime | Select-Object -First 1

if($null -eq $cabFile){ 
    throw "No .cab files found in current directory : $workingDirectory"
}

#construct full path uppon file were extracted by cab filename
$cabExtractedDirectory = '{0}\{1}' -f $tempDirectory, $cabFile.Basename.Split('-')[0]

#try to create temp location to extract cab file
New-Item -Path $tempDirectory -ItemType Directory -ErrorAction SilentlyContinue

#if this cab file was already extracted, try to delete old extraction directory
Remove-Item -Path $cabExtractedDirectory -Recurse -Force -ErrorAction SilentlyContinue

#extrat cab file in temporary location
cmd.exe /c C:\Windows\System32\expand.exe -F:* "$workingDirectory\$($cabFile.Name)" $tempDirectory

#iterate all .inf files and try to install them
Get-ChildItem -Path $cabExtractedDirectory -Include *.inf -Recurse | ForEach-Object { pnputil.exe /Add-driver $_.FullName /Install }

#return to orignal script call location
Set-Location $originalWorkingDirectory

Write-Host "Installation success." -ForegroundColor Green -BackgroundColor Black
pause
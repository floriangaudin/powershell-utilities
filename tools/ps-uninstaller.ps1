<#
    .SYNOPSIS
    Audit or uninstall application

    .DESCRIPTION
    The ps-uninstaller.ps1 allow audit or execution
    of uninstallation command

    .PARAMETER Product
    Specifies the product name

    .PARAMETER Version
    Optionnal. Specifies the version of the product

    .PARAMETER Audit
    Specifies the mode of execution. By default is true
    to prevent any command to run and execute.

    .OUTPUTS
    Generates host generics outputs

    .EXAMPLE
    PS> .\ps-uninstaller.ps1 -Product Adobe

    .EXAMPLE
    PS> .\ps-uninstaller.ps1 -Product Adobe -Version "22.001.20117"

    .EXAMPLE
    PS> .\ps-uninstaller.ps1 -Product "Adobe Acrobat DC (64-bit)" -Audit:$false
#>

Param(
    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Product,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$Version,
    [parameter()]
    [ValidateNotNullOrEmpty()]
    [switch]$Audit=$true)

#get default sysdir
function Get-Platform{
    if($Env:PROCESSOR_ARCHITECTURE -eq "AMD64"){
        return "$env:Systemroot\SysWOW64"  
    }elseif($Env:PROCESSOR_ARCHITECTURE -eq "x86"){
        return = "$env:Systemroot\System32"  
    }
}

#get a desired product from the registry
function Find-ProductsFromRegistry{
    Param(
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Product,
        [parameter(Mandatory=$false)]
        [string]$Version
    )
    #registry entries to search for
    $paths  = @(
        "HKLM:\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", 
        "HKLM:\SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        ) 

    #array contains any registry entries
    $entries = @()

    #iterate any registry entries and add them in array
    foreach ($path in $paths) {
        Get-ChildItem -Path $path | ForEach-Object {

            $entries += Get-ItemProperty $_.pspath      
        }
    }

    #gets products found with the criteria -Product
    $products = $entries | Where-Object {$_.DisplayName -like "*$Product*"} | Select-Object -Unique DisplayName, DisplayVersion, UninstallString | Sort-Object DisplayName

    #more of one products found and the criteria -Version is not present
    if(($products | Measure-Object).Count -gt 1 -and $Version -eq ""){

        Write-Host "`n"
        Write-Host "Showing results with the criteria `"" -ForegroundColor Red -NoNewline
        Write-Host "$Product" -ForegroundColor Green -BackgroundColor Black -NoNewline
        Write-Host "`"" -ForegroundColor Red
        Write-Host "(NOTE : specify -Product or/and -Version to select the desired product)" -ForegroundColor DarkYellow

        #display list of products found
        $products | Select-Object @{Name="Product";Expression={$_.DisplayName}}, @{Name="Vesrion";Expression={$_.DisplayVersion}} | Format-Table -AutoSize
        break
        
    }
    #more of one products found and the criteria -Version is present
    elseif(($products | Measure-Object).Count -gt 1 -and $Version -ne ""){

        #get the targeted application by the -Version criteria
        return $products | Where-Object { $_.DisplayVersion -like "*$Version*"}

    #only one product found
    }else{

        return $products

    }
}

#gets command to uninstall product found within registry entries
function Get-UninstallCommand {
    param(
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$UninstallString
    )

    
    #REGEX indicates if UninstallString contains an msi commands 
    $MsiRegex = '^(?:msiexec|msiexec.exe).*?(/[IX]).*?({\w{8}-\w{4}-\w{4}-\w{4}-\w{12}})'

    #REGEX indicates if UninstallString does not contains an msi commands, otherwise, this is an exe command
    $ExeRegex = '^(?!msiexec|msiexec.exe)(".*?.exe"|.*?.exe)(.*?$)'

    #available exe switches
    $ExeSwitches = "/S /V /qn /ms /uninstall /quiet /VERYSILENT /NORESTART"

    #available msi switches
    $MsiSwitches = "/qn /quiet /norestart"

    $MSImatch = $UninstallString -match $MsiRegex
    $EXEmatch = $UninstallString -match $ExeRegex

    #UninstallString contains an msi command
    if($MSImatch -eq $true){

        #split UninstallString with regex by 3 match's groups : command, arg, guid
        $split = $UninstallString -split $MsiRegex
     
        #possible values $split[0] : msiexec, msiexec.exe
        #possible values $split[1] : /i, /I, /x, /X
        #possible values $split[2] : {00000000-1111-2222-3333-444444444444}, {AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}

        #return full-formatted msi command 
        return @("$(Get-Platform)\msiexec.exe", "$($split[1]) $($split[2]) $MsiSwitches")
    }

    #UninstallString contains an exe command
    if($EXEmatch -eq $true){

        #split UninstallString with regex by 3 match's groups : command, arg, guid
        $split = $UninstallString -split $ExeRegex
        
        #value $split[0] : null (negative regex condition (?!msiexec|msiexec.exe) : does not start with ...)
        #possible values $split[1] : absolute path to executable uninstaller program
        #possible values $split[2] : can be null or contains executable uninstaller program args

        #if any, replace "" char by '' to prevent any space in absolute path
        if(-not $split[1].StartsWith('"'))
        {  
            $split[1] = "`"$($split[1])`"" 
        }

        #return full-formatted exe command 
        return @("$(Get-Platform)\cmd.exe","/c $($split[1]) $($split[2]) $ExeSwitches")

    }

}

#display an alert if -Audit:$true (true by default)
if($Audit.IsPresent) {
    Write-Host "                                                                      " -ForegroundColor Yellow -BackgroundColor Gray
    Write-Host " AUDIT " -ForegroundColor Yellow -BackgroundColor Gray -NoNewline
    Write-Host "`t`tSet " -ForegroundColor Gray -NoNewline
    Write-Host "-Audit:`$false " -ForegroundColor Green -NoNewline
    Write-Host "to allow script to run`t" -ForegroundColor Gray -NoNewline
    Write-Host " MODE " -ForegroundColor Yellow -BackgroundColor Gray
    Write-Host "                                                                      " -ForegroundColor Yellow -BackgroundColor Gray

}

$products = Find-ProductsFromRegistry -Product $Product -Version $Version

#no application found, abort
if(-not $products){
    break
}

#products found in registry entries within criterias
if($null -ne $products -and ($products | Measure-Object).Count -gt 0){

    #iterate products
    foreach($p in $products){

        #get command for product
        $command = Get-UninstallCommand -UninstallString $p.UninstallString

        #display command
        Write-Host "[$($p.DisplayName):$($p.DisplayVersion)] Uninstall Command : " -ForegroundColor Yellow -NoNewline
        Write-Host $command -ForegroundColor DarkYellow
    
        #uninstall product only if -Audit is set to false
        if(-not $Audit.IsPresent){

            #run default uninstall command
            Start-Process -FilePath $command[0] -ArgumentList $command[1] -NoNewWindow -Wait

            #usefull in case of msiexec uninstall command, /i (msiexec installation switch) can be found in registry UninstallString and is considered as uninstallation switch on registry...
            #if /i does not work, try a second one with inverted switches, rerun command with replacement of /i by /x (Microsoft mindblowing....)
            Start-Process -FilePath $command[0] -ArgumentList ($command[1] -replace '/i', '/x' -replace '/I','/X') -NoNewWindow -Wait

            Write-Host "$($p.DisplayName) ($($p.DisplayVersion)) successfully uninstalled." -ForegroundColor Green
        }

    }

#no product found in registry entries, trying to uninstall product directly via Get-Package cmdlet...
}else{
    Get-Package -Provider Programs | Where-Object {$_.Name -like "*$Product*"} | ForEach-Object { 
        if(-not $Audit.IsPresent){
            $_.Uninstall() 
        }
    } -ErrorAction Stop

}
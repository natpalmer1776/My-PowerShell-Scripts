########################################
## PLACE DEVELOPER DOCUMENTATION HERE ##
## 
## Dependency [Hash-String.ps1]
##    Purpose: Converts the supplied string to a hash using the specified hashing algorithm
## 
## Function []
##    Purpose: 
##    Calls: 
## 
## Global Variable []
##    Purpose: 
## 
## PLACE DEVELOPER DOCUMENTATION HERE ##
########################################

##############################
## IMPORT DEPENDENCIES HERE ##

. "$PSScriptRoot\Log-Core.ps1"

## IMPORT DEPENDENCIES HERE ##
##############################

#######################
## DECLARE VARIABLES ##

Set-Variable -Name "computer" -Scope "script"
Set-Variable -Name "computerAD" -Scope "script"
Set-Variable -Name "computerEnclosure" -Scope "script"
Set-Variable -Name "computerOSDetails" -Scope "script"
Set-Variable -Name "computerOSName" -Scope "script"
Set-Variable -Name "computerBIOS" -Scope "script"

## DECLARE VARIABLES ##
#######################

#######################################
## DEFINE APPLICATION FUNCTIONS HERE ##

function Log-Computer {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$computer,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$path,
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]$application
    )

    $script:computer = $computer
    
    if(Test-Connection -ComputerName $computer -Quiet -TimeToLive 5) {
        Gather-Info
        $payload = @{ TypeID = "Log-Computer"; "Computer Name" = $computer; "Manufacturer" = Query-Manufacturer; "Model" = Query-Model;
            "Serial" = Query-Serial; "Chassis" = Query-Chassis; "OS" = Query-OperatingSystem;
            "Department" = Query-Department; "Facility" = Query-Facility; "Last User" = Query-LastUser;
            "Last Logon Date" = Query-LastLogon; "Uptime" = Query-UpTime }
        if($application){$payload.Add("Application", $application)}
        if($path.EndsWith(".json")) { $logObj = New-Log -Path $path -Data $payload } else { $logObj = New-Log -Path ($path + "\$computer.json") -Data $payload }
    } else {Write-Host "$computer is offline or unreachable."}

}

function Gather-Info {
    $script:computerAD = Get-ADComputer -identity $script:computer -Properties *
    $script:computerEnclosure = Get-WmiObject -ComputerName $script:computer -Class Win32_SystemEnclosure
    $script:computerOS = Get-WmiObject -ComputerName $script:computer -Class Win32_OperatingSystem
    $script:computerBIOS = Get-WmiObject -ComputerName $script:computer -Class Win32_BIOS
    $script:computerSystem = Get-WmiObject -ComputerName $script:computer -Class Win32_ComputerSystem
    $script:computerLogon = Get-WmiObject -ComputerName $script:computer -Class Win32_NetworkLoginProfile
}

function Query-Manufacturer {

    $manufacturer = $script:computerBIOS.Manufacturer

    return [String]$manufacturer
}

function Query-Model {

    $model = $script:computerSystem.Model

    return [String]$model
}

function Query-Serial {

    $serial = $script:computerBIOS.SerialNumber

    return [String]$serial
}

function Query-Chassis
{

    # List of possible chassis types to be returned from WMI and their string counterparts
    $enclosure_list = @{
        1="Other";
        2="Unknown";
        3="Desktop ";
        4="Low Profile Desktop";
        5="Pizza Box";
        6="Mini Tower";
        7="Tower";
        8="Portable";
        9="Laptop";
        10="Notebook";
        11="Hand Held";
        12="Docking Station";
        13="All in One";
        14="Sub-Notebook";
        15="Mini PC";
        16="Lunch Box";
        17="Main System Chassis";
        18="Expansion Chassis";
        19="SubChassis";
        20="Bus Expansion Chassis";
        21="Peripheral Chassis";
        22="RAID Chassis";
        23="Rack Mount Chassis";
        24="Sealed-case PC";
        25="Multi-system chassis";
        26="Compact PCI";
        27="Advanced TCA";
        28="Blade";
        29="Blade Enclosure";
        30="Tablet";
        31="Convertible";
        32="Detachable";
        33="IoT Gateway";
        34="Embedded PC";
        35="Mini PC";
        36="Stick PC"
    }
    return [String]($enclosure_list.[Int]($computerEnclosure.ChassisTypes[0]))
}

function Query-OperatingSystem {

    if($script:computerOS.Name.Indexof("|")){$operatingsystem = $script:computerOS.Name.SubString(0,$script:computerOS.Name.IndexOf("|"))}
    else{$operatingsystem = $script:computerOS.Name}

    return [String]$operatingsystem
}

function Query-Department {

    if($script:computer.Length -eq 12){
        $department = $script:computer.SubString(7,2)
    }

    # TODO: Implement user configurable dept code (IS, ED, RG, etc.) dictionary to display full department name.

    return [String]$department
}

function Query-Facility {
    $computer = $script:computer

    $computerDN = ([adsisearcher]"(&(objectCategory=computer)(objectClass=computer)(cn=$computer))").FindOne().Properties.distinguishedname
    $parentOU = $computerDN.SubString(("$computerDN".IndexOf("DC"))-6,5)
    $facility = (([adsisearcher]"(&(objectCategory=organizationalUnit)(name=$parentOU))").FindOne().Properties.description)
    if($facility){if("$facility".IndexOf("(")){$facility = $facility.SubString(0, ("$facility".IndexOf("("))-1)}}

    return [String]$facility
}

function Query-LastUser {

    $user = $script:computerLogon | Sort-Object -Property LastLogon -Descending | 
        Select-Object -Property * -First 1 | 
        Where-Object {$_.LastLogon -match "(\d{14})"} | % {$_.Name}

    if($user){return [String]$user}else{return "None"}
}

function Query-LastLogon {

    $lastLogon = $script:computerLogon | Sort-Object -Property LastLogon -Descending | 
        Select-Object -Property * -First 1 | 
        Where-Object {$_.LastLogon -match "(\d{14})"} | % {[datetime]::ParseExact($matches[0], "yyyyMMddHHmmss", $null)}

    if($lastlogon){return [String]($lastlogon.ToString())}else{return "None"}
}

function Query-UpTime {

    $uptime = ($computerOS.ConvertToDateTime($computerOS.LocalDateTime))-($computerOS.ConvertToDateTime($computerOS.LastBootUpTime))
    if($uptime.Days -gt 1){$grammarD = " Days, "} else {$grammarD = " Day, "}
    if($uptime.Hours -gt 1){$grammarH = " Hours, "} else {$grammarH = " Hour, "}
    if($uptime.Minutes -gt 1){$grammarM = " Minutes"} else {$grammarM = "Minute"}
    $uptime = $uptime.Days.ToString()+$grammarD+$uptime.Hours+$grammarH+$uptime.Minutes+$grammarM

    return [String]$uptime
}

## DEFINE APPLICATION FUNCTIONS HERE ##
#######################################
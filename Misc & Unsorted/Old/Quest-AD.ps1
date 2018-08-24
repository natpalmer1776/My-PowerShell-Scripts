#Required Snapin for Active Directory Scripts - Needs Active Roles Management Shell for Active Directory Installed [.msi]
Add-PSSnapin Quest.ActiveRoles.ADManagement
Write-Host "Starting EasyAD Script:"
if(-not $global:ADSession)
{
    Write-Host "Initializing Active Directory Session..."
    $global:ADSession = Connect-QADService -proxy
    Write-Host "Initialization complete."
}
else
{
    Write-Host ("    Using current Active Directory session " + $global:ADSession.DefaultNamingContext.Path)
    Write-Host "Initialization complete."    
}

#Creates a reference to an AD Object and stores it in a variable. Works for Users, Computers, etc.
$user = Get-QADUser "mcoagenrespiratory"

# Queries Active-Directory for the computer object under the specified name and returns it with full properties.
function EAD-Computer
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$computer
    )

    return Get-QADComputer -Name $computer -Connection $global:ADSession 
}

function EAD-User
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$user
    )
    
    return (Get-QADUser -Identity "$user" -Connection $global:ADSession)
}

function EAD-UserLastLogon
{
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [String]$user
    )
    return [datetime]::fromfiletime((query-AD $user).lastLogon)
}

function EAD-ComputerLastLogon
{
    param
    (
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$computer
    )

    $computer_ad = EAD-Computer $computer
    return (Get-Date $computer_ad.Item("lastLogonTimeStamp"))
}

function EAD-ComputerFacility
{
    param
    (
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$Computer
    )
    # Get the distinguished name of the local computer from active directory.
    $computerDN = ([adsisearcher]"(&(objectCategory=computer)(objectClass=computer)(cn=$Computer))").FindOne().Properties.distinguishedname
    $domainIndex = "$computerDN".IndexOf("DC")
    $parentOU = $computerDN.SubString($domainIndex-6,5)
    $facilityDescription = ([adsisearcher]"(&(objectCategory=organizationalUnit)(name=$parentOU))").FindOne().Properties.description
    return $facilityDescription
}
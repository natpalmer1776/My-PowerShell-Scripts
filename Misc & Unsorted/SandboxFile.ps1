clear-host

# Gets the list of computers from the text file
$computers = gc .\computers.txt
$prevKey

# Add the following Registration Key
$Path1 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths"
$Property1 = "\\*\NETLOGON"
$Value1 = "RequireMutualAuthentication=1, RequireIntegrity=1"

# Add the following Registration Key
$Path2 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths"
$Property2 = "\\*\SYSVOL"
$Value2 = "RequireMutualAuthentication=1, RequireIntegrity=1"

# Processes each line in computers.txt
$results = foreach ($computer in $Computers)
{

# Test to see if the workstation is online
If (test-connection -ComputerName $computer -Count 1 -Quiet)
{

# If workstation is online, adds the new registry settings, reports output as "Success"
Try {
    Invoke-Command -Computer $computer {Get-Childitem HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider} | % {$prevKey = $_.Property}

    $status = "Success"
    write-host $computer+" completed."

    # If workstation is online, but did not apply the patch, reports output of "Failed"
    } Catch {
        write-host $computer+" failed."
        $status = "Failed"
    }

    $postKey
    Invoke-Command -Computer $computer {Get-Childitem HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider} | % {$postKey = $_.Property}
}

# If workstation is offline, reports output of "Workstation Unreachable"
else
{
    $status = "Workstation Unreachable"
}

$i = $postKey.Count
$current2 = 1
$finalString2 = ""

While($current2 -le $i) {
    $finalString2 += ("(("+$postKey[$current-1]+"))")
    write-host "test"
    $current++
}

# Creates the CSV file with the headers of "Computer" and "Status"
New-Object -TypeName PSObject -Property @{
    'Post Key' = $postKey
    'Computer'=$computer
    'Status'=$status
    'Previous Key'=$finalString
}
$finalString=$null
$preVkey=$null
$postKey=$null
}

# Outputs the status of each workstation progress
$results |
Export-Csv -NoTypeInformation -Path "./AddRegKey.csv"
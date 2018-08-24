$computers = New-Object System.Collections.ArrayList($null)
$names = New-Object System.Collections.ArrayList($null)
$global:connected = New-Object System.Collections.ArrayList($null)
$totaljobs = 0

Get-ADComputer -Filter * -SearchBase "OU=WS00015,OU=00015,DC=hca,DC=corpad,DC=net" | % {$computers.Add($_)}
$computers | % {$names.Add($_.Name)}

$script = {
    Param($computer)
    $global:totaljobs++
    if(Test-Connection -Quiet -Count 2 $computer) {$global:connected.Add($computer)}
}

$names | % {Start-Job -ScriptBlock $script -ArgumentList $_}
Write-Host ("There are "+$totaljobs+" computers being checked.")
do {
    $running = (Get-Job -Stating Running).Count
    Write-Progress -id 0 -PercentComplete ($running / $totaljobs) -Activity "Checking computers for connectivity"
    Sleep -Seconds 1
} while(Get-Job -State 'Running')

Write-Progress -id 0 -complete

$payload = @($computers, $names, $connected)
return $payload
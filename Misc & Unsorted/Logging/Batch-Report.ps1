. "$PSScriptRoot\Log-Computer.ps1"
. "$PSScriptRoot\Log-Core.ps1"

$ListPath = 'C:\users\fvo7197\downloads\Computer Logs\computers.txt'
$LogPath = 'C:\users\fvo7197\downloads\Computer Logs\'

Get-Content $ListPath | % {Log-Computer -computer $_ -path $LogPath -application "Batch Report"}
Generate-Report -Source "C:\users\fvo7197\downloads\Computer Logs\" -path "C:\report.xlsx"
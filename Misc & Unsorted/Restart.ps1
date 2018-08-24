param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$ComputerName
)

Invoke-WebRequest -Uri "http://live.sysinternals.com/psexec.exe" -OutFile "$Env:tmp\psexec.exe"


if(Test-Path $("filesystem::\\$ComputerName\c$\"))
{
    . "$ENV:TMP\psexec.exe" \\$ComputerName /accepteula cmd /c "shutdown /r /t 0"
    Add-Content -Path "C:\restartingcomputers.txt" -Value $ComputerName
}
else
{
    Add-Content -Path "C:\offlinecomputers.txt" -Value $ComputerName
}
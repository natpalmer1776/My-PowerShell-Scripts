function Install-Printers {
    param(
        [String]$ComputerName,
        [String]$Icon
    )
    if(Test-Connection -ComputerName $ComputerName -Count 2 -Quiet) {
        Copy-Item -Path C:\printers.txt -Destination "\\$ComputerName\C$\" -Force
        Copy-Item -Path C:\AddPrinters.bat -Destination "\\$ComputerName\C$\" -Force
        PSEXEC \\$ComputerName "C:\AddPrinters.bat"
    }

}
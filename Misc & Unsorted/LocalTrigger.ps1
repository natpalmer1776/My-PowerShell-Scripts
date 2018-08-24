cd C:\

.\PsExec.exe \\MCOAH1NISC12.hca.corpad.net -u "HCA\FVO7197" -p "Lexie96$" -i /accepteula cmd /c "powershell -windowstyle hidden -noninteractive -executionpolicy unrestricted -file C:\script.ps1"
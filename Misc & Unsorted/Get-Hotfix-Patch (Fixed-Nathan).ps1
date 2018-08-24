# Set script execution preferences
    $ErrorActionPreference = “SilentlyContinue”

# Spreadsheet Location & Filename
    $DirectoryToSaveTo = "c:\scripts\checkpatch\" 
    $date=Get-Date
    $Filename="Patchinfo-$($date)" 
  
# Read computers.txt file contents
    $Computers = Get-Content "c:\scripts\checkpatch\computers.txt" 

# Enter KB to be checked here 
    $Patch = Read-Host 'Enter KB'
  
# Test File path & access, create directory if it doesn't exist
    if (!(Test-Path -path "$DirectoryToSaveTo")) { New-Item "$DirectoryToSaveTo" -type directory | out-null } 

# Create a new Excel object using COM  
    $Excel = New-Object -ComObject Excel.Application 
    $Excel.visible = $True 
    $Excel = $Excel.Workbooks.Add() 
    $Sheet = $Excel.Worksheets.Item(1) 
    
# Create a Title for the first worksheet 
    $Sheet.Name = 'Patch status - '

# Set the currently active Row & Cell
    $row = 1 
    $Column = 1 

# Set the value of the first cell to "Patch Status"
    $Sheet.Cells.Item($row,$column)= 'Patch status'

# Format the sheet between A1 & F2
    $range = $Sheet.Range("A1","F2") 
    $range.Merge() | Out-Null 
    $range.VerticalAlignment = -4160 
 
#Give it a nice Style so it stands out 
    $range.Style = 'Title' 
 
#Increment row for next set of data 
    $row += 2
 
# Save the initial row so it can be used later to create a border 
# Counter variable for rows
    $intRow = $row 
    $xlOpenXMLWorkbook=[int]51 
 
# Set the header for each column
    $Sheet.Cells.Item($intRow,1)  ="Name" 
    $Sheet.Cells.Item($intRow,2)  ="status" 
    $Sheet.Cells.Item($intRow,3)  ="Patch status" 
    $Sheet.Cells.Item($intRow,4)  ="OS" 
    $Sheet.Cells.Item($intRow,5)  ="SystemType" 
    $Sheet.Cells.Item($intRow,6)  ="Last Boot Time" 

# Formats the column headers
    for ($col = 1; $col –le 6; $col++) { 
          $Sheet.Cells.Item($intRow,$col).Font.Bold = $True 
          $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48 
          $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 34 }

# Select the next row after finishing the headers
    $intRow++
 
# Returns the string equivalent of the supplied Status Code
Function GetStatusCode 
{  
    Param([int] $StatusCode)   
    switch($StatusCode) 
    { 
        0         {"Success"} 
        11001   {"Buffer Too Small"} 
        11002   {"Destination Net Unreachable"} 
        11003   {"Destination Host Unreachable"} 
        11004   {"Destination Protocol Unreachable"} 
        11005   {"Destination Port Unreachable"} 
        11006   {"No Resources"} 
        11007   {"Bad Option"} 
        11008   {"Hardware Error"} 
        11009   {"Packet Too Big"} 
        11010   {"Request Timed Out"} 
        11011   {"Bad Request"} 
        11012   {"Bad Route"} 
        11013   {"TimeToLive Expired Transit"} 
        11014   {"TimeToLive Expired Reassembly"} 
        11015   {"Parameter Problem"} 
        11016   {"Source Quench"} 
        11017   {"Option Too Big"} 
        11018   {"Bad Destination"} 
        11032   {"Negotiating IPSEC"} 
        11050   {"General Failure"} 
        default {"Failed"} 
    } 
} 
 
 
# Take the raw date/time string & format it to a user-friendly format
Function GetUpTime 
{ 
    param([string] $LastBootTime) 
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime) 
    "Days: $($Uptime.Days); Hours: $($Uptime.Hours); Minutes: $($Uptime.Minutes); Seconds: $($Uptime.Seconds)"  
} 
 
# Process each computer in computers.txt
    foreach ($Computer in $Computers) { 
        $error.clear()
        write-host ("Connecting to " + $Computer)
        TRY {  
            $system = Invoke-Command -ComputerName $Computer -ScriptBlock {
                [Hashtable]$results
                $results = @{
                    $OS = (Get-WmiObject -Class Win32_OperatingSystem);
                    $sheetS = (Get-WmiObject -Class Win32_ComputerSystem);
                    $sheetPU = (Get-WmiObject -Class Win32_Processor);
                    $drives = (Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3});
                    $pingStatus = (Get-WmiObject -Query "Select * from win32_PingStatus where Address='$Computer'");
                    $uptime = ((Get-WmiObject win32_operatingsystem | select csname, @{LABEL='LastBootUpTime'; EXPRESSION={ $_.ConverttoDateTime($_.lastbootuptime) } } ).LastBootUpTime);
                    $OSRunning = ($OS.caption + " " + $OS.OSArchitecture + " SP " + $OS.ServicePackMajorVersion);
                    $systemType = ($sheetS.SystemType);
                    $date = (Get-Date)}

                $results
            }
   
            if ($kb=get-hotfix -id $Patch -ComputerName $computer -ErrorAction 2) { $kbinstall="$patch is installed" } 
            else { $kbinstall="$patch is not installed" } 

            if($pingStatus.StatusCode -eq 0) { $Status = GetStatusCode( $system.pingStatus.StatusCode ) } 
            else { $Status = GetStatusCode( $system.pingStatus.StatusCode ) }
            write-host ($Computer + " completed")} 
        
        CATCH { 
        write-host $error
        $pcnotfound = $True}
 
# Sent Data to Excel 
    if ($pcnotfound -eq $True) { 
        #$sheet.Cells.Item($intRow, 1) = "PC Not Found" 
        $sheet.Cells.Item($intRow, 1) = $computer 
        $sheet.Cells.Item($intRow, 2) = "PC Not Found" } 
    else { 
        $sheet.Cells.Item($intRow, 1) = $computer 
        $sheet.Cells.Item($intRow, 2) = $status 
        $sheet.Cells.Item($intRow, 3) = $kbinstall 
        $sheet.Cells.Item($intRow, 4) = $system.OSRunning 
        $sheet.Cells.Item($intRow, 5) = $system.sheetS.SystemType 
        $sheet.Cells.Item($intRow, 6) = $system.uptime } 
 
        $intRow = $intRow + 1 
        $pcnotfound = "false" } 
 
# Autofit Excel Document Columns
    $Sheet.UsedRange.EntireColumn.AutoFit() 

########################################
 
########################################

    $filename = "$DirectoryToSaveTo$filename.xlsx" 
# if (test-path $filename ) { rm $filename } #delete the file if it already exists 
    $Sheet.UsedRange.EntireColumn.AutoFit() 
    $Excel.SaveAs($filename, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx) 
    $Excel.Saved = $True 
    #$Excel.Close() 
    #$Excel.DisplayAlerts = $False 
    #$Excel.quit()

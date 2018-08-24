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

## 
## CHANGE PATH PARAMETER TO ACCEPT WITHOUT .JSON & APPEND LOGID.JSON TO PATH STRINGS WITHOUT.
## 

##############################
## IMPORT DEPENDENCIES HERE ##



## IMPORT DEPENDENCIES HERE ##
##############################

#######################################
## DEFINE APPLICATION FUNCTIONS HERE ##

function New-Log {
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({$_.EndsWith(".json")})]
        [String]$Path,

        [Parameter(Mandatory=$false)]
        [Hashtable]$Data
    )

    if($Data){[Hashtable]$Log = $Data.Clone()}
    else{$Log = [Hashtable]::New()}

    $date = Get-Date -DisplayHint Date

    # Create a Log Key
    $key = New-Object System.Collections.ArrayList($null)
    if($Log.GetEnumerator().Name.Count -gt 0){
        $Log.GetEnumerator().Name | % {
            if($_ -ne "TypeID"){$key.Add($_) | out-null}
        }
    }

    # Add the hashtable key to the Hashtable
    $Log.Add("Key",$key)

    # Add the current date to the Hashtable
    $Log.Add("Date",$date)

    # Adds the current path to the Hashtable
    $Log.Add("Path",$Path)

    # Verify/Set Log TypeID
    if( -not $Log.TypeID) {
        $Log.Add("TypeID","Default")
    }

    # Verify/Generate LogID
    if( -not $Log.ID) {
        $logID = Get-Random -Minimum 1000000000 -Maximum 9999999999
        $Log.Add("ID",$logID)
    }

    # Create a custom powershell object with the log & path as properties
    $logObj = New-Object PSCustomObject -Property $Log

    $Log = @{}

    #########################################
    ## DEFINE CUSTOM OBJECT FUNCTIONS HERE ##

        <# Add-Member -MemberType ScriptMethod -InputObject $logObj -Name "Sample" -Value { function goes here } #>

        Add-Member -MemberType ScriptMethod -InputObject $logObj -Name "Save" -Value {
            param
            (
                [Parameter(Mandatory=$false)]
                [String] $path
            )

            if(!$path.EndsWith(".json")) {
                write-host "Unable to save log, please ensure a filename & path ending in .json is specified in the path parameter"
            } else{$this.Path = $path}

            $this | ConvertTo-Json -Depth 32 -Compress | Set-Content -Path $this.Path
        }

        Add-Member -MemberType ScriptMethod -InputObject $logObj -Name "Reload" -Value {
            param
            (
                [Parameter(Mandatory=$false)]
                [String] $NewPath
            )

            if($NewPath -and $NewPath.EndsWith(".json")) {
                $this.Path = $NewPath
            }
            elseif(-not (Test-Path $this.Path)){
                write-host $this.Path
                write-host "Supplied path is invalid, using predefined Path attribute."
            }

            $newObj = Get-Content -Path $this.Path | ConvertFrom-Json
            $newObj.PSObject.Properties | % {
                $name = $_.Name
                $value = $_.Value

                if($this.$name) {
                    if($name -ne "Key") {
                        $this.$name = $value
                    }
                }
            }            
        }

    ## DEFINE CUSTOM OBJECT FUNCTIONS HERE ##
    #########################################

    #Returns the PsCustomObject "$logObj"
    $logObj.Save($Path)
    return $logObj
}

function Load-Log {
    param
    (
    [Parameter(Mandatory=$true)]
    [ValidateScript({$_.EndsWith(".json")})]
    [String]
    $Path
    )

    if((Test-Path $Path)) {
        return (Get-Content -Path $Path | ConvertFrom-Json)
    }
}

##                            ##
#  CREATE GENERATE-REPORT GUI  #
##                            ##

function Generate-Report {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$source,
        [parameter(Mandatory=$false, Position=2)]
        [ValidateScript({$_.EndsWith(".xlsx")})]
        [String]$path,
        [parameter(Mandatory=$false, Position=2)]
        [ValidateNotNullOrEmpty()]
        [String]$TypeID,
        [parameter(Mandatory=$false, Position=1)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$Filter
    )

    ## FILTER FORMAT #############
    #    Key | Val    | Require? #
    ##############################
    # TypeID | String | Required #
    #  Field | String | Optional #
    # S.Date | Date   | Optional #
    # E.Date | Date   | Optional #
    ##############################

    [Boolean]$use_sdate = $false
    [Boolean]$use_edate = $false

    if($Filter) {
        if(-not $Filter.TypeID){$Filter.Add("TypeID", "Default")}
        if($Filter.sdate){$use_sdate = $true}
        if($Filter.edate){$use_edate = $true}
    }

    $Excel = New-Object -ComObject Excel.Application 
    $Excel.visible = $True
    $Workbook = $Excel.Workbooks.Add() 
    $Sheet = $Workbook.Worksheets.Item(1)
    $Sheet.Name = "Log Report"

    $filelist = New-Object System.Collections.ArrayList($null)
    $files = New-Object System.Collections.ArrayList($null)
    $files_filter = New-Object System.Collections.ArrayList($null)
    $files_final = New-Object System.Collections.ArrayList($null)
    Get-ChildItem $source | % {if($_.FullName.EndsWith(".json")){$filelist.Add($_.FullName) | out-null}}
    $filelist | % {$files.Add((Load-Log $_)) | out-null}
    
    if($use_sdate){$files | % {if($_.Date -lt $Filter.sdate){$files_filter.Add($_) | out-null}}}
    if($use_edate){$files | % {if($_.Date -gt $Filter.edate){$files_filter.Add($_) | out-null}}}
    $files | % {if($_.TypeID -ne $Filter.TypeID){$files_filter.Add($_) | out-null}}

    if($Filter) {
        $Filter.GetEnumerator() | % {
            if($_.Name -ne "sdate" -and $_.Name -ne "edate" -and $_.Name -ne "TypeID" -and $_.Name -ne "Path")
            {
                $compare_name = $_.Name
                $compare_value = $_.Value
                $files | % {
                    if($_.$compare_name){
                        if(-not (($_.$compare_name) -like $compare_value))
                        {
                            $files_filter.Add($_) | out-null
                        }
                    } else {$files_filter.Add($_) | out-null}
                }
            }
        }
    $files | % {if(-not ($files_filter -contains $_)){$files_final.Add($_) | out-null}}
    } else { $files | % {$files_final.Add($_)} | out-null }

    $total_columns = 2
    $startcol = 3
    $current_row = 2
    $current_column = 1

    $Sheet.Cells.Item(1,1) = "ID"
    $Sheet.Cells.Item(1,2) = "Date"

    if($files_final[0].Key.Contains("Application")) {
        $Sheet.Cells.Item(1,3) = "Application"
        $startcol++
        $total_columns++
    }

    $files_final[0].Key | % {
        $index = $files_final[0].Key.IndexOf($_)
        $column = $index + $startcol
        $Sheet.Cells.Item(1,$column) = $_
        $total_columns++
    }

    $files_final | % {
        while($current_column -le $total_columns)
        {
            $selected_property = $Sheet.Cells.Item(1, $current_column).Formula
            $Sheet.Cells.Item($current_row, $current_column) = ($_.$selected_property)
            $current_column++
        }
        $current_row++
        $current_column = 1
    }

    $Sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    if(-not $path) {
        $timestamp = Get-Date -Format FileDateTime
        $Workbook.SaveAs("$source\$timestamp.xlsx",51)
    } else {$Workbook.SaveAs("$path",51)}
    #$Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel
}
########################################
## PLACE DEVELOPER DOCUMENTATION HERE ##
## 
## Dependency []
##    Purpose: 
## 
## Function []
##    Purpose: 
##    Calls: 
## 
## Static Variable []
##    Purpose: 
## 
## PLACE DEVELOPER DOCUMENTATION HERE ##
########################################

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Add-Type -AssemblyName PresentationCore,PresentationFramework

function New-Form {
    param
    (
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="String")]
        [ValidateNotNullOrEmpty()]
        [String]$source,
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="Path")]
        [ValidateScript({($_.EndsWith(".xaml") -and (Test-Path $_))})]
        [String]$path
    )

    if($path){$source = Get-Content $path -Raw}
    $source = $source -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$xaml = $source
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    try{$form=[Windows.Markup.XamlReader]::Load($reader)}
    catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
    $SyncHash = @{}
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object{
        Set-Variable -Name "$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        $SyncHash.Add($_.Name, (Get-Variable -Name $_.Name))
    }
    #$Form.ShowDialog()

    return $form
}

function New-Runspace {
    param
    (
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="String")]
        [ValidateNotNullOrEmpty()]
        [String]$source,
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="Path")]
        [ValidateScript({($_.EndsWith(".xaml") -and (Test-Path $_))})]
        [String]$path
    )

    $syncHash = [HashTable]::Synchronized(@{})
    $newRunspace = [runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $psCmd = [PowerShell]::Create().AddScript({
        param
        (
            [Parameter(Mandatory=$true, Position=0, ParameterSetName="String")]
            [ValidateNotNullOrEmpty()]
            [String]$source,
            [Parameter(Mandatory=$true, Position=0, ParameterSetName="Path")]
            [ValidateScript({($_.EndsWith(".xaml") -and (Test-Path $_))})]
            [String]$path
        )
        if($path){$source = Get-Content $path -Raw}
        $source = $source -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

        [xml]$xaml = $source
        $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window=[Windows.Markup.XamlReader]::Load($reader)
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {$syncHash.Add($_.Name, $syncHash.Window.FindName($_.Name))}
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error = $Error
    })

    if($source){$psCmd.AddParameter("source", $source)}else{$psCmd.AddParameter("path", $path)}

    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()

    return $syncHash
}
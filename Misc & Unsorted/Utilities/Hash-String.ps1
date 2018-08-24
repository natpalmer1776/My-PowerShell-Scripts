function Hash()
{
    param
    (
        [parameter(Mandatory = $true, ParameterSetName="MD5", Position=0)]
        [Switch]$MD5,
        [parameter(Mandatory = $true, ParameterSetName="RIPEMD160", Position=0)]
        [Switch]$RIPEMD160,
        [parameter(Mandatory = $true, ParameterSetName="SHA512", Position=0)]
        [Switch]$SHA512,
        [parameter(Mandatory = $true, ParameterSetName="SHA384", Position=0)]
        [Switch]$SHA384,
        [parameter(Mandatory = $true, ParameterSetName="SHA256", Position=0)]
        [Switch]$SHA256,
        [parameter(Mandatory = $true, ParameterSetName="SHA1", Position=0)]
        [Switch]$SHA1,

        [parameter(Mandatory = $true, ParameterSetName="MD5", Position=1)]
        [parameter(Mandatory = $true, ParameterSetName="RIPEMD160", Position=1)]
        [parameter(Mandatory = $true, ParameterSetName="SHA512", Position=1)]
        [parameter(Mandatory = $true, ParameterSetName="SHA384", Position=1)]
        [parameter(Mandatory = $true, ParameterSetName="SHA256", Position=1)]
        [parameter(Mandatory = $true, ParameterSetName="SHA1", Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]$Value
    )

    Add-Type -assemblyName "System.Security"

    [String]$hash = $null
    [String]$key_string = $null
    [byte[]]$to_hash = $null
    [byte[]]$byte_array = $null

    if($MD5){
        $hasher = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
    } elseif($RIPEMD160){
        $hasher = New-Object System.Security.Cryptography.RIPEMD160Managed
    } elseif($SHA512){
        $hasher = New-Object System.Security.Cryptography.SHA512Managed
    } elseif($SHA384){
        $hasher = New-Object System.Security.Cryptography.SHA384Managed
    } elseif($SHA256){
        $hasher = New-Object System.Security.Cryptography.SHA256Managed
    } elseif($SHA1){
        $hasher = New-Object System.Security.Cryptography.SHA1Managed
    }
    
    $to_hash = [System.Text.Encoding]::UTF8.GetBytes($Value)
    $byte_array = $hasher.ComputeHash($to_hash)
    foreach($byte in $byte_array)
    {
        $hash += $byte.ToString()
    }

    return $hash
}
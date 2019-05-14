<#
    .SYNOPSIS
        Relocates the Outlook OST file by rebuilding it after ForceOSTPath registry key change.

    .DESCRIPTION
        After changing the registry key ForceOSTPath (either using a GPO or by direct registry 
        editing) the OST file is not relocated by Outlook. This script will force Outlook to rebuild
        the OST in the new location. Preferably run this script at logon.

    .PARAMETER LogFile
        Write script actions the specified log file.
        Note:
        If ommitted no log file will be written and output will go to console only. If a relative path is 
        given the log file will be written to the ForceOSTPath location.

    .OUTPUTS
        A boolean is returned.
        $true if a relocation action has been initiated by updating the registry keys.
        $false if no relocation has been initiated. In other words, the OST does not need to be relocated.

    .EXAMPLE
        Powershell.exe -NoLogo -NonInteractive -File Relocate-OutlookOST.ps1

    .LINK
        Links to further documentation.

    .NOTES
        The scripts has a few limitations, pending features and assumptions:
        - Tested and developed on Windows 2016/10 with Outlook 2016/365
        - Only the default Outlook profile is relocated
        - Relocation is done by rebuilding the OST and not by moving the file
        - No other files are relocated
        - The script assumes Outlook is installed and does not test for that
#>
Param(
    [Parameter(Mandatory=$false, Position=0)]
    [String]
    $LogFile
)

# Base location of the Outlook registry key
$OutlookRegistryBase = "HKCU:\Software\Microsoft\Office\16.0\Outlook"

# Get Outlook GPO settings
$OutlookPolicies = Get-ItemProperty "HKCU:\Software\Policies\Microsoft\office\16.0\outlook" -ErrorAction SilentlyContinue

# Some logic that allows local testing without GPO set
if ($null -eq $OutlookPolicies.forceostpath) {
    Write-Error "Registry key ForceOSTPath not set"
    return
} else {
    $OSTPath = $OutlookPolicies.forceostpath
    # Create the path if it doesn't exist
    New-Item $OSTPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
}

# Building the log file location
if ($LogFile -ne "") {
    if (! $(Split-Path $LogFile -IsAbsolute)) {
        $LogFile = Join-Path $OSTPath $LogFile
    }
    $WriteLog = $true
} else {
    $WriteLog = $false
}

function Cleanup-Log {
    <#
    .SYNOPSIS
        Reduced the log file if needed
    #>
    param (
        [Parameter(Mandatory=$false, Position=0)]
        [String]
        $File,
        [Parameter(Mandatory=$false, Position=1)]
        [Int]
        $MaxSize
    )
    $Content = Get-Content $File
    if ($($Content.count) -gt $MaxSize) {
        $Content | Select-Object -Skip $($($Content.count)-$MaxSize) | Set-Content $File
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes output to a log file and console.
    .DESCRIPTION
        Write-Log will write to a file and console if $write is set to $true.
        Other wise only outputs to screen
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false, Position=0)]
        [String]
        $Msg,
        [Parameter(Mandatory=$true, Position=1)]
        [boolean]
        $Write,
        [Parameter(Mandatory=$false, Position=2)]
        [String]
        $File
    )
    $LogDate = Get-Date -Format "yyyy-MMM-dd"
    $LogTime = Get-Date -Format "HH:mm:ss"
    $Level = "INFO"
    # Date Time - Host - Level - Message
    $LogLine = "$($LogDate) $($LogTime) - $($ENV:COMPUTERNAME) - $($Level) - $($Msg)"
    if ($write) {
        Add-content $File -value $LogLine
        Write-Host $LogLine
        Cleanup-Log $File 500
    } else {
        Write-Host $LogLine
    }
}

function Convert-REGBinaryToString {
    <#
    .SYNOPSIS
        Converts a REG_BINARY value to a some what readable string.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [System.Byte[]] $BinaryRegValue
    )
    [System.Text.Encoding]::Default.GetString($BinaryRegValue) -replace "`0",""
}

function Convert-StringToREGBinary {
    <#
    .SYNOPSIS
        Converts a string to a REG_BINARY value.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $StringRegValue
    )
    $ReturnStr = ''
    for ($i=0; $i -lt $StringRegValue.Length; $i++) {
        $ReturnStr += $newpath[$i]
        $ReturnStr += [char]0
    }
    $ReturnStr | Format-Hex
}

Write-Log "Start of the relocation script." $WriteLog $LogFile
Write-Log "ForceOSTPath value: $($OSTPath)" $WriteLog $LogFile

# Get all the Outlook registry keys
$OutlookSettingsRoot = Get-ItemProperty $OutlookRegistryBase

# Find default Outlook profile registry keys
$OutlooDefaultProfileBase = "$OutlookRegistryBase\Profiles\$($OutlookSettingsRoot.DefaultProfile)\"
$OutlookProfile = Get-ChildItem $OutlooDefaultProfileBase
Write-Log "Outlook default profile name: $($OutlookSettingsRoot.DefaultProfile)" $WriteLog $LogFile

# Gotcha is the array with keys that are of intrest to us.
# Keys of interest are keys that have a value that matches the name of the default profile
$gotcha = @()
# Loop through the keys
foreach ($RegKey in $($OutlookProfile)) {
    # Get the subkeys with values
    $KeyDetails = Get-Item "$OutlooDefaultProfileBase\$($RegKey.PSChildName)"
    # Loop through the values
    foreach ($Value in $($KeyDetails.GetValueNames())) {
        # Skip value if it iss binary for script performance
        if ($RegKey.GetValueKind($value) -ne "Binary") {
            # if the value matches the default profile name we have a match.
            if ($KeyDetails.GetValue($Value) -eq $($OutlookSettingsRoot.DefaultProfile)) {
                # Write-Host "Found a profile key"
                $gotcha += $KeyDetails.PSChildName
                Write-Log "  Found key of interrest: $($KeyDetails.PSChildName)" $WriteLog $LogFile
            }
        }
    }
}
Write-Log "Total number of keys of interrest: $($gotcha.count)" $WriteLog $LogFile

# Loop through the keys of interest
foreach ($FoundKey in $gotcha) {
    # Get the subkeys with values
    $FoundKeyDetails = Get-Item "$OutlooDefaultProfileBase\$FoundKey"
    # Loop through the values
    Write-Log "Processing key: $($FoundKey)" $WriteLog $LogFile
    foreach ($Value in $($FoundKeyDetails.GetValueNames())) {
        # Now we do want to process the Binary values
        if ($FoundKeyDetails.GetValueKind($Value) -eq "Binary") {
            # Convert the binary value to something more readable
            $strValue = Convert-REGBinaryToString($FoundKeyDetails.GetValue($Value))
            # Check if the value has .ost in it indicating it's the path to the file
            if ($strValue -imatch ".ost") {
                Write-Log "  Registry key: $Value" $WriteLog $LogFile
                Write-Log "  Current OST path: $strValue" $WriteLog $LogFile
                # Extract the name of the ost file
                $OSTfileName = $strValue.split('\')[-1]
                Write-Log "  OST file name: $OSTfileName" $WriteLog $LogFile
                # Build the new full path
                $newpath = "$OSTPath\$OSTfileName"
                Write-Log "  New OST location: $newpath" $WriteLog $LogFile
                # If the old and new paths mismatch it must be a moved
                if ($strValue -ieq $newpath) {
                    Write-Log "!!! OST allready at location specified in ForceOSTPath !!!" $WriteLog $LogFile
                    $false
                } else {
                    # Convert the path to binary registry value
                    $RegBinary = Convert-StringToREGBinary($newpath)
                    # Set the new location in the registry
                    Write-Log "!!! Forcing Outlook to relocate OST on next start !!!" $WriteLog $LogFile
                    Set-ItemProperty -Path $FoundKeyDetails.PSpath -Name $Value -Value $RegBinary.Bytes -verbose -ErrorAction 'Stop'
                    $true
                }
            }
        }
    }
}

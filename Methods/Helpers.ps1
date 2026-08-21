function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[$Level] $Message"
    Add-Content -Path $LogPath -Value $line
    Write-Host $line
}

function New-RandomPassword {
    param([int]$Length = 16)
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%'
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

function Escape-LdapFilterValue {
    param([string]$Value)
    return $Value.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([char]0, '\00')
}

function Get-ExistingAdUserForHire {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$SamAccountName
    )

    if (-not [string]::IsNullOrWhiteSpace($FirstName)) { $FirstName = $FirstName.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($LastName)) { $LastName = $LastName.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($SamAccountName)) { $SamAccountName = $SamAccountName.Trim() }

    $props = 'Department', 'EmployeeID', 'Description', 'DistinguishedName', 'SamAccountName', 'GivenName', 'Surname', 'Name'

    if (-not [string]::IsNullOrWhiteSpace($SamAccountName)) {
        $bySam = Get-ADUser -Filter "SamAccountName -eq '$(Escape-LdapFilterValue $SamAccountName)'" `
            -Properties $props -ErrorAction SilentlyContinue
        if ($bySam) { return @($bySam)[0] }
    }

    if (-not [string]::IsNullOrWhiteSpace($FirstName) -and -not [string]::IsNullOrWhiteSpace($LastName)) {
        $byName = Get-ADUser -Filter "GivenName -eq '$(Escape-LdapFilterValue $FirstName)' -and Surname -eq '$(Escape-LdapFilterValue $LastName)'" `
            -Properties $props -ErrorAction SilentlyContinue
        if ($byName) { return @($byName)[0] }

        $displayName = "$FirstName $LastName"
        $byDisplayName = Get-ADUser -Filter "Name -eq '$(Escape-LdapFilterValue $displayName)'" `
            -Properties $props -ErrorAction SilentlyContinue
        if ($byDisplayName) { return @($byDisplayName)[0] }
    }

    return $null
}

function Test-RowStatusShouldSkip {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }

    $skipValues = @('Processed', 'Staged', 'Failed', 'Error', 'Skipped')
    return [bool]($skipValues | Where-Object { $_ -ieq $Value.Trim() })
}

function Mark-HireRowSkipped {
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][int]$Index,
        [Parameter(Mandatory)][string]$SamAccountName,
        [string]$Reason = 'Account already exists in AD'
    )

    Set-RowProperty -Row $Rows[$Index] -Name 'Processed' -Value 'Skipped'
    Set-RowProperty -Row $Rows[$Index] -Name 'Username' -Value $SamAccountName
    Set-RowProperty -Row $Rows[$Index] -Name 'SkipReason' -Value $Reason
    Save-SpreadsheetRows -Rows $Rows
    Write-Log "Marked row $Index ($SamAccountName) as Skipped"
}

function New-Username {
    param([string]$First, [string]$Last)
    $base = ("$($First.Substring(0,1))$Last" -replace '[^a-zA-Z]', '').ToLower()
    $candidate = $base
    $i = 1
    while (Get-ADUser -Filter "SamAccountName -eq '$candidate'" -ErrorAction SilentlyContinue) {
        $i++
        $candidate = "$base$i"
    }
    return $candidate
}

function Get-RowField {
    param(
        $Row,
        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return [string]$prop.Value.Trim()
        }
    }

    return $null
}

function Set-RowProperty {
    param(
        $Row,
        [Parameter(Mandatory)][string]$Name,
        $Value
    )

    $Row | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}

function Test-RowShouldSkip {
    param($Row)

    $skipValues = @('Processed', 'Staged', 'Failed', 'Error', 'Skipped')
    $processed = Get-RowField -Row $Row -Names @('Processed')
    $status    = Get-RowField -Row $Row -Names @('Status')
    if (Test-RowStatusShouldSkip -Value $processed) { return $true }
    if (Test-RowStatusShouldSkip -Value $status) { return $true }

    return $false
}

function Test-RowIsExample {
    param($Row)

    $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
    $notes     = Get-RowField -Row $Row -Names @('Notes', 'Note')

    if ($firstName -eq "First" -and $lastName -eq "Last") { return $true }
    if ($notes -match 'example row') { return $true }

    return $false
}

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
    if ($processed -in $skipValues) { return $true }
    if ($status -in $skipValues) { return $true }

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

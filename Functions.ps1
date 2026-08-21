<#
.SYNOPSIS
    Loader for all provisioning modules. Dot-sourced by Main.ps1.

.DESCRIPTION
    Loads config and job functions from Private/. Does not run on its own.
    Run Main.ps1 to execute the provisioning loop.
#>

$PrivateRoot = Join-Path $PSScriptRoot 'Methods'

. (Join-Path $PrivateRoot 'Config.ps1')
. (Join-Path $PrivateRoot 'Helpers.ps1')
. (Join-Path $PrivateRoot 'Spreadsheet.ps1')
. (Join-Path $PrivateRoot 'Get-PendingNewHires.ps1')
. (Join-Path $PrivateRoot 'Invoke-StageUser.ps1')
. (Join-Path $PrivateRoot 'EmployeeId.ps1')
. (Join-Path $PrivateRoot 'Invoke-SortUser.ps1')
. (Join-Path $PrivateRoot 'Invoke-ReturnToStaging.ps1')
. (Join-Path $PrivateRoot 'Invoke-ProvisionNewHire.ps1')

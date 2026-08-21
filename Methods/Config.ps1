# ============================================================
# CONFIG - edit these for your environment
# ============================================================

# Pick ONE source mode: Local | Url | SharePoint
$SpreadsheetMode   = "Local"

# Local mode - file path on this machine or UNC share
$SpreadsheetPath    = "C:\HR_Drive\HR_NewHires.xlsx"

# Url mode - direct HTTPS link that downloads the .xlsx file
$SpreadsheetUrl     = "https://yourtenant.sharepoint.com/.../HR_NewHires.xlsx"
$SpreadsheetUploadUrl = ""

# SharePoint mode - recommended for production
$SharePointSiteUrl          = "https://yourtenant.sharepoint.com/sites/HR"
$SharePointServerRelativeUrl = "/sites/HR/Shared Documents/HR_NewHires.xlsx"
$SharePointConnectParams = @{ Interactive = $true }

# Local working copy - always used for Import-Excel / Export-Excel
$LocalWorkbookPath   = "C:\ProvisioningLogs\HR_NewHires.xlsx"
$WorksheetName       = "Active"
$SpreadsheetStartRow = 4
$LogPath             = "C:\ProvisioningLogs\provisioning_$(Get-Date -Format 'yyyyMMdd').log"
$DefaultDomain       = "mydomain.com"
$DefaultPassLen      = 16

$StagingOU      = "OU=1NewUserStaging,OU=Active-Users,DC=mydomain,DC=com"
$EmployeeIDFile = "C:\ProvisioningLogs\next_employee_id.txt"
$EmployeeIDPrefix = "EMP"
$EmployeeIDSeed = 1001

$SmtpServer = "smtp.yourdomain.com"
$MailFrom   = "ad-automation@yourdomain.com"
$MailTo     = "it-team@yourdomain.com"

# Department -> OU + Groups mapping. Groups can stay @() until real group names are confirmed.
$DeptMap = @{
    "RK Industries"     = @{ OU = "OU=RK Industries,OU=Active-Users,DC=mydomain,DC=com";     Groups = @() }
    "RK Mechanical"     = @{ OU = "OU=RK Mechanical,OU=Active-Users,DC=mydomain,DC=com";     Groups = @() }
    "RK Electrical"     = @{ OU = "OU=RK Electrical,OU=Active-Users,DC=mydomain,DC=com";     Groups = @() }
    "RK Steel"          = @{ OU = "OU=RK Steel,OU=Active-Users,DC=mydomain,DC=com";          Groups = @() }
    "RK Water"          = @{ OU = "OU=RK Water,OU=Active-Users,DC=mydomain,DC=com";          Groups = @() }
    "RK Energy"         = @{ OU = "OU=RK Energy,OU=Active-Users,DC=mydomain,DC=com";         Groups = @() }
    "RK Service"        = @{ OU = "OU=RK Service,OU=Active-Users,DC=mydomain,DC=com";        Groups = @() }
    "RK Mission Critical" = @{ OU = "OU=RK Mission Critical,OU=Active-Users,DC=mydomain,DC=com"; Groups = @() }
}

Import-Module ActiveDirectory
Import-Module ImportExcel -ErrorAction Stop

$Script:WorkbookPath = $LocalWorkbookPath
$Script:SharePointConnected = $false

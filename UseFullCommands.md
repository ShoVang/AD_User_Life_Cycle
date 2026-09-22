# Import the AD module if not already loaded

Import-Module ActiveDirectory

# Set the SamAccountName(s) you want to KEEP

$Keep = @("sho.vang", "administrator", "svc-adprovisioning") # adjust to match your actual admin/service accounts

# Preview what WOULD be deleted (safe - no changes made)

Get-ADUser -Filter \* -SearchBase "OU=Active-Users,DC=mydomain,DC=com" -Properties SamAccountName |
Where-Object { $Keep -notcontains $\_.SamAccountName } |
Select-Object SamAccountName, Name, DistinguishedName

    Get-ADUser -Filter * -SearchBase "OU=Active-Users,DC=mydomain,DC=com" -Properties SamAccountName |
    Where-Object { $Keep -notcontains $_.SamAccountName } |
    Remove-ADUser -Confirm:$true

###################################

# cd C:\Scripts\AD_User_Life_Cycle

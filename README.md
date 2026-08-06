# AD User Lifecycle Provisioning

Automates new-hire AD account creation from an HR spreadsheet: stages a bare account, assigns a sequential EmployeeID, sorts it into the correct department OU, and adds group memberships — one user at a time, per pipeline run.

## How it works

Each pending row in the HR spreadsheet goes through this pipeline in `Invoke-ProvisionNewHire`:

1. **Grab** — `Get-PendingNewHires` picks the next row where `Processed` isn't `"Processed"` and `FirstName`/`LastName` aren't blank.
2. **Stage** — `Invoke-StageUser` creates the AD account in the staging OU (or locates it if already staged).
3. **Assign ID** — `Invoke-AssignEmployeeId` pulls the next sequential ID from a local counter file (`EMP1001`, `EMP1002`, ...).
4. **Sort** — `Invoke-SortUser` moves the account to its department OU and adds it to the mapped groups, based on `$DeptMap`.
5. **Return to staging** — if the department is unmapped or the sort fails, the account stays in (or goes back to) the staging OU and its EmployeeID is cleared.

`Main.ps1` dot-sources `Functions.ps1` and loops, re-reading the spreadsheet after every user, until no pending rows remain. It's a straight run-through — there's no `-Phase` flag; staging and sorting happen back-to-back per user, not in separate batch passes.

## Prerequisites

- **PowerShell modules:** `ActiveDirectory`, `ImportExcel`
  ```powershell
  Install-Module ImportExcel -Scope CurrentUser -Force
  ```
- **Folders that must exist before running** (the script does NOT create these automatically):
  ```powershell
  New-Item -ItemType Directory -Path C:\ProvisioningLogs -Force
  ```
- **AD rights:** whatever account runs this needs delegated rights on the staging OU and every target department OU (create/move/modify users, add to groups).
- **Network:** if installing `ImportExcel` for the first time, the VM needs outbound access to PowerShell Gallery.

## Configuration (top of `Functions.ps1`)

| Variable                                | Purpose                                                                |
| --------------------------------------- | ---------------------------------------------------------------------- |
| `$SpreadsheetMode`                      | `Local`, `Url`, or `SharePoint` — source for the HR spreadsheet        |
| `$SpreadsheetPath`                      | Local/UNC path to the source `.xlsx` (Local mode)                      |
| `$LocalWorkbookPath`                    | Working copy the script actually reads/writes                          |
| `$WorksheetName`                        | Must match the real tab name in the workbook **exactly**               |
| `$LogPath`                              | Daily log file, `C:\ProvisioningLogs\provisioning_<date>.log`          |
| `$StagingOU`                            | DN of the staging OU                                                   |
| `$EmployeeIDFile`                       | Local counter file tracking the next EmployeeID                        |
| `$DeptMap`                              | Hashtable mapping spreadsheet `Department` values → target OU + groups |
| `$SmtpServer` / `$MailFrom` / `$MailTo` | Summary email after each run                                           |

## Known environment-specific values (this lab)

These were confirmed against the real domain during testing — make sure `Functions.ps1` actually matches:

- Domain: `mydomain.com`
- Real OU structure is flat under `Active-Users`, **not** under a `CORP` or generic `Users` parent:
  - `OU=1NewUserStaging,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Electrical,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Energy,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Industries,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Mechanical,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Mission Critical,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Service,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Steel,OU=Active-Users,DC=mydomain,DC=com`
  - `OU=RK Water,OU=Active-Users,DC=mydomain,DC=com`
  - No dedicated `IT` OU exists yet — decide where IT hires should land before enabling that department in `$DeptMap`.
- HR spreadsheet lives at `C:\HR_Drive\HR_NewHires.xlsx`.
- **`$DeptMap` still only has placeholder entries** (`Sales`, `Accounting`, `IT`) — needs real entries for every `RK *` department above, with correct group names, before a live run will sort anyone correctly.

> Run `Get-ADOrganizationalUnit -Filter * | Select Name, DistinguishedName` any time to re-confirm the current OU list if the AD structure changes.

## Running it

From an elevated PowerShell console, as an account with rights to the staging and department OUs:

```powershell
cd C:\Scripts\AD_User_Life_Cycle
.\Main.ps1
```

Output streams to the console and to `C:\ProvisioningLogs\provisioning_<date>.log`.

## Verifying a run

Check who's sitting in staging (nothing here after a successful run means everyone sorted correctly):

```powershell
Get-ADUser -Filter * -SearchBase "OU=1NewUserStaging,OU=Active-Users,DC=mydomain,DC=com" -Properties Department, EmployeeID | Select Name, SamAccountName, Department, EmployeeID
```

Spot-check a specific hire:

```powershell
Get-ADUser -Filter "GivenName -eq 'First' -and Surname -eq 'Last'" -Properties Department, EmployeeID
```

Inspect what the script actually sees in the spreadsheet before troubleshooting "no pending hires":

```powershell
Import-Excel -Path "C:\HR_Drive\HR_NewHires.xlsx" -WorksheetName "<worksheet name>" | Select -First 3 | Format-List
```

## Troubleshooting log (issues hit so far in this lab)

| Symptom                                   | Cause                                                                          | Fix                                                             |
| ----------------------------------------- | ------------------------------------------------------------------------------ | --------------------------------------------------------------- |
| `Functions.ps1` not recognized            | Wrong path/filename                                                            | Confirm exact filename with `Get-ChildItem`                     |
| `ImportExcel` module not loaded           | Not installed on this machine                                                  | `Install-Module ImportExcel -Scope CurrentUser -Force`          |
| `Add-Content` directory not found         | `C:\ProvisioningLogs` doesn't exist                                            | `New-Item -ItemType Directory -Path C:\ProvisioningLogs -Force` |
| Spreadsheet not found                     | `$SpreadsheetPath` doesn't match real filename                                 | `Get-ChildItem C:\HR_Drive`, fix path or rename file            |
| Worksheet not found                       | `$WorksheetName` doesn't match real tab name                                   | Match exactly (typos and all) or rename the tab in Excel        |
| "No pending new hires found"              | `Processed`/`FirstName`/`LastName` columns missing or already marked Processed | `Import-Excel ... \| Format-List` to inspect actual row data    |
| `Get-ADUser` "Directory object not found" | OU DN in a command/script doesn't match real AD structure                      | `Get-ADOrganizationalUnit -Filter *` to get real DNs            |

## Before a live/production run

- [ ] Spreadsheet has a `Processed` column (blank for new rows)
- [ ] `$DeptMap` has an entry for every real department value used in the sheet
- [ ] `$StagingOU` matches the real staging OU DN
- [ ] `C:\ProvisioningLogs` exists
- [ ] SMTP settings are either real and tested, or the `Send-MailMessage` block is disabled for early runs
- [ ] Tested end-to-end against a 1-2 row copy of the spreadsheet first
- [ ] Edge cases tested: unmapped department, missing manager, already-processed row, duplicate name/username collision

## Not yet handled

- No mechanism currently surfaces the generated initial password to anyone (it's created and used, then discarded) — decide how new hires/IT will actually receive credentials.
- No `IT` OU exists in current AD structure despite `IT` being a valid department value in the spreadsheet.

How Does It Work — code call order

SETUP (happens once when Main.ps1 starts, before anything runs)
1   You run .\Main.ps1 on the server
2   Main.ps1 calls Functions.ps1 — Functions.ps1 loads every file in Methods/ so all the methods exist
3   Config.ps1 runs first — sets paths (HR_Drive, ProvisioningLogs), domain, staging OU, department map, and loads the AD + Excel modules
4   Helpers.ps1 loads — logging, usernames, passwords, reading/writing row fields
5   Spreadsheet.ps1 loads — all the Excel read/write methods
6   The rest load — Get-PendingNewHires, Invoke-StageUser, EmployeeId, Invoke-SortUser, Invoke-ReturnToStaging, Invoke-ProvisionNewHire
7   Main.ps1 calls the Main method

MAIN RUN (what actually runs each time)
8   Write-Log — logs that the run started
9   Import-SpreadsheetFromSource (Spreadsheet.ps1) — copies HR_NewHires.xlsx from C:\HR_Drive to C:\ProvisioningLogs (does not read row data yet, just gets the file)

LOOP (repeats until no pending hires left)
10  Import-SpreadsheetRows (Spreadsheet.ps1) — opens the file, reads the Active tab starting at row 4, gives back all rows as objects (FirstName, LastName, Department, etc.)
11  Get-PendingNewHires (Get-PendingNewHires.ps1) — loops through those rows and keeps only ones that still need work (blank Processed, not the example row, has a first and last name)
12  If nothing pending — loop stops, go to step 25
13  Main picks the first pending hire and calls Invoke-ProvisionNewHire (Invoke-ProvisionNewHire.ps1)

INSIDE Invoke-ProvisionNewHire (one hire at a time)
14  Invoke-StageUser (Invoke-StageUser.ps1) — checks if account already exists in AD; if yes, marks row Skipped and stops this hire
15  If no account exists — New-ADUser creates them in OU=1NewUserStaging
16  Save-SpreadsheetRows (Spreadsheet.ps1) — writes Processed = "Staged" back to the Excel file and copies it back to HR_Drive
17  Resolve-DepartmentMapping (Invoke-SortUser.ps1) — looks up their department in $DeptMap (e.g. "RK Industries" → that OU path)
18  If department not mapped — user stays in staging, loop goes back to step 10 for the next hire
19  Invoke-AssignEmployeeId (EmployeeId.ps1) — assigns the next EmployeeID (EMP1001, EMP1002, …) from the counter file
20  Invoke-SortUser (Invoke-SortUser.ps1) — Move-ADObject moves the user from staging into their department OU
21  If sort fails — Invoke-ReturnToStaging (Invoke-ReturnToStaging.ps1) moves them back to 1NewUserStaging, loop goes back to step 10
22  Save-SpreadsheetRows — writes Processed = "Processed", EmployeeID, and date back to Excel
23  Loop goes back to step 10 — re-reads the spreadsheet and picks the next pending hire

END OF RUN
24  If any hires succeeded — Send-MailMessage sends a summary email to IT
25  Write-Log — logs that the run finished

Spreadsheet Processed column after a run
    blank     = not touched yet
    Staged    = account created, still in 1NewUserStaging
    Processed = done, moved to department OU
    Skipped   = account already existed in AD
    Failed    = error — check ErrorMessage column

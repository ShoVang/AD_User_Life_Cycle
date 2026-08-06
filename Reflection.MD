What I Learned
A reflection on building an Active Directory home lab and automating new-hire provisioning with PowerShell — from spinning up VMs to debugging a real script against a real domain.

Setting up the VMs
I started with Oracle VirtualBox and built two machines: a Windows Server box to act as the domain controller, and a Windows 11 client to represent a normal end-user machine. Keeping them separate forced me to actually think about client/server relationships instead of just clicking around one box — the DC does the domain work, the client is what a real employee's machine would look like joined to that domain.

One of the earliest and most important lessons had nothing to do with PowerShell at all: domain naming has to be consistent everywhere, from day one. I originally set things up expecting DC=mydomain,DC=local and it turned out the real domain was DC=mydomain,DC=com. That mismatch quietly broke things downstream until I got in the habit of never assuming — always running Get-ADDomain or Get-ADOrganizationalUnit -Filter \* to check the real structure before writing anything that referenced it.

Setting up administrators and structure
I built out the OU structure by hand before automating anything — Users, Computers, Groups, service accounts, and department OUs. Doing this manually first mattered: it meant that when I later wrote scripts that referenced these paths, I actually understood what a Distinguished Name was and why the path had to be exact, instead of just copying syntax I didn't understand.

I also created a dedicated service account (svc-adprovisioning) instead of running automation as a personal admin account — with PasswordNeverExpires and CannotChangePassword set deliberately, since a scheduled task can't type in a new password when one expires. That was a small decision but it taught me to think about automation accounts differently from human accounts from the start.

Shared folders and groups
Setting up HR_Drive and ADScripts as SMB shares taught me an ordering lesson the hard way: the AD objects a share command references have to already exist before you reference them. Trying to set permissions using a group or user that didn't exist yet in AD threw a Windows error I hadn't seen before (System Error 1332) and forced me to slow down and build things in the right order — create the identity first, then grant it access, not the other way around.

Keeping naming consistent across the folder name, the share name, and any related AD group also turned out to matter more than I expected. Early on when names drifted slightly between layers, it made troubleshooting permissions confusing later, so I made a habit of keeping them aligned.

Learning PowerShell
This project is where PowerShell actually clicked for me, mostly because I wasn't learning it in the abstract — I was learning it to solve a real, annoying problem: turning a three-step manual process (HR spreadsheet → manager ticket → manual AD setup) into one automated pipeline.

Splitting the code into Functions.ps1 (a dot-sourced library of reusable functions) and Main.ps1 (the entry point that actually runs the pipeline) taught me a real separation-of-concerns habit instead of writing one giant script. Debugging real errors — missing modules, wrong file paths, mismatched worksheet names, wrong OU paths — taught me more about how PowerShell actually behaves than any tutorial did, because each error had a specific, traceable cause I had to hunt down myself:

Modules like ImportExcel don't just exist — they have to be installed on the specific machine running the script.
Cmdlets like Add-Content don't create missing parent folders; you have to do that yourself first.
A worksheet name in code has to match the real tab name in Excel character-for-character, typos included.
A hardcoded OU path in a script is only as good as your last Get-ADOrganizationalUnit check — AD structure drifts, and code doesn't know that unless you update it.
My process
Looking back at how I actually worked through this, a pattern shows up: run it, read the error, fix exactly what the error says, run it again. Not guessing ahead, not trying to fix five things at once — one error at a time, in the order they actually appeared. That kept me from chasing phantom problems and helped me actually understand each failure instead of just making it go away.

I also learned to test small before testing big — using a trimmed-down copy of the spreadsheet instead of running the full HR list against a script I hadn't fully verified yet, so that if something went wrong, it went wrong on two fake rows instead of twenty real ones.

The other piece of the process was git — cloning the repo onto the server, treating the script as source-controlled code rather than a one-off file I edited directly on the VM. That habit is going to matter more once this moves toward being scheduled and running unattended.

What's next
The lab and script work isn't done — I still need to finish mapping every real department OU, double check the staging OU path, decide how a new hire's initial password actually gets to them, and eventually get this running unattended through Task Scheduler. But the debugging process so far taught me the actual mechanics of AD, PowerShell, and Windows Server in a way that just reading documentation never would have.

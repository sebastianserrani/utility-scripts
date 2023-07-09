# Utility scripts

A collection of utility scripts.

## Powershell
* `Add-ActiveDirectoryUserInteractively.ps1` Interactively create an Active Directory User account
* `Import-ActiveDirectoryOU.ps1` Import Active Directory OUs from a CSV file
* `Import-ActiveDirectoryUser.ps1` Imports Active Directory User accounts from a CSV file
* `List-LocalUSerPasswordExpiration.ps1` Check LOCAL users for upcoming password expiration
* `Notify-LocalUSerPasswordExpiration.ps1` Notify LOCAL users for upcoming password expiration
* `Reset-ActiveDirectoryUserPasswordExpiration.ps1` Checks if Active Directory users have expired passwords

## VBScript
* `AddLocalUser.vbs` Adds a new local user
* `MozillaThunderbirdClean.vbs` Deletes *.mozmsgs folders, and *.mozeml and *.wdseml files within
* ### Functions
    * `CmdOutput.vbs` Returns the StdOut from a console command
    * `CPULoad.vbs` Returns the CPU load
    * `DateFileCreated.vbs` Returns the creation date of a file
    * `DateFileLastMod.vbs` Returns the last modification date of a file
    * `DateISO.vbs` Returns the date in ISO 8601 format (YYYY-MM-DD)
    * `DateTimeAmerican.vbs` Returns the date in american format (MM-DD-YYYY HH:MM:SS)
    * `DateTimestamp.vbs` Returns the seconds from Unix Epoch on January 1st, 1970
    * `DownloadFile.vbs` Downloads a file from an URL
    * `ElevatePermissions.vbs` Relaunches the calling script asking for elevated permissions
    * `ErrorType.vbs` Returns the type of an error (Error, Warning, Information)
    * `FileCopy.vbs` Copies a file
    * `FileMove.vbs` Moves a file
    * `ForceCScriptExecution.vbs` Forces the script to run in a command-line environment
    * `HTTPPost.vbs` POST to an URL endpoint
    * `IniFileRead.vbs` Reads an .ini files and returns a dictionary of values
    * `Log.vbs` Logs to a .log file with the name of the script
    * `MemoryFreePhysical.vbs` Returns the total free memory of the computer in MB
    * `MemoryTotal.vbs` Returns the total memory of the computer in MB
    * `ProgressMsg.vbs` Displays a non-blocking message box
    * `RndPassword.vbs` Generates a semi random password
    * `TimeSpan.vbs` Returns the time span in hours:minutes:seconds format
    * 

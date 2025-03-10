# Microsoft 365 Auditor script

<a target="_blank" href="https://github.com/ivancarlosti/m365auditor"><img src="https://img.shields.io/github/stars/ivancarlosti/m365auditor?style=flat" /></a>
<a target="_blank" href="https://github.com/ivancarlosti/m365auditor"><img src="https://img.shields.io/github/last-commit/ivancarlosti/m365auditor" /></a>
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors)](https://github.com/sponsors/ivancarlosti)

This script collects users, groups, Teams of a [Microsoft 365 Business]([https://workspace.google.com/](https://www.microsoft.com/microsoft-365/business)) environment on .xlsx file for audit and review purposes, the file is archived in a .zip file including a screenshot with hash MD5 of the .xlsx file and the script executed.

## Instructions

* Save .ps1, .txt files locally (download [here](https://github.com/ivancarlosti/m365auditor/zipball/master))
* Change variables of `mainscript.ps1` if needed
* Run `mainscript.ps1` on PowerShell (right-click on file > Run with PowerShell)
* Follow instructions selecting project name, option 1 to generate audit report and collect .zip file on `$destinationpath`

## Requirements

* Windows 10+ or Windows Server 2019+
* PowerShell
* Modules `AzureAD`, `MicrosoftTeams`, `ImportExcel` on PowerShell

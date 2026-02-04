# Microsoft 365 Auditor script
This script collects users, groups and Teams of a Microsoft 365 environment on .xlsx file for audit and review purposes

<!-- buttons -->

<!-- endbuttons -->

## Instructions
* Save the last release version and extract files locally (download [here](https://github.com/ivancarlosti/m365auditor/releases/latest))
* Change variables of `mainscript.ps1` if needed
* Update `tenantIds.txt` with your tenants
* Run `mainscript.ps1` on PowerShell (right-click on file > Run with PowerShell)
* Follow instructions selecting tenant, authenticate, collect .zip file on `$destinationpath`
* If needs help to install or update required modules, run `ADMIN-install-modules.ps1` as administrator

## Requirements
* Windows 10+ or Windows Server 2019+
* PowerShell 5.x (some modules still runs under .NET Framework, PowerShell 7.x uses .NET Core)
* Modules `MicrosoftTeams`, `ImportExcel`, `Microsoft.Graph`, `ImportExcel` on PowerShell

<!-- footer -->
---

## 🧑‍💻 Consulting and technical support
* For personal support and queries, please submit a new issue to have it addressed.
* For commercial related questions, please [**contact me**][ivancarlos] for consulting costs. 

| 🩷 Project support |
| :---: |
If you found this project helpful, consider [**buying me a coffee**][buymeacoffee]
|Thanks for your support, it is much appreciated!|

[ivancarlos]: https://ivancarlos.me
[buymeacoffee]: https://www.buymeacoffee.com/ivancarlos

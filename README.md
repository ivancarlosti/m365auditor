# Microsoft 365 Auditor script
This script collects users, groups and Teams of a Microsoft 365 environment on .xlsx file for audit and review purposes

<!-- buttons -->
[![Stars](https://img.shields.io/github/stars/ivancarlosti/awssesconverter?label=⭐%20Stars&color=gold&style=flat)](https://github.com/ivancarlosti/awssesconverter/stargazers)
[![Watchers](https://img.shields.io/github/watchers/ivancarlosti/awssesconverter?label=Watchers&style=flat&color=red)](https://github.com/sponsors/ivancarlosti)
[![Forks](https://img.shields.io/github/forks/ivancarlosti/awssesconverter?label=Forks&style=flat&color=ff69b4)](https://github.com/sponsors/ivancarlosti)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/ivancarlosti/awssesconverter?label=Activity)](https://github.com/ivancarlosti/awssesconverter/pulse)
[![GitHub Issues](https://img.shields.io/github/issues/ivancarlosti/awssesconverter?label=Issues&color=orange)](https://github.com/ivancarlosti/awssesconverter/issues)
[![License](https://img.shields.io/github/license/ivancarlosti/awssesconverter?label=License)](LICENSE)  
[![GitHub last commit](https://img.shields.io/github/last-commit/ivancarlosti/awssesconverter?label=Last%20Commit)](https://github.com/ivancarlosti/awssesconverter/commits)
[![Security](https://img.shields.io/badge/Security-View%20Here-purple)](https://github.com/ivancarlosti/awssesconverter/security)
[![Code of Conduct](https://img.shields.io/badge/Code%20of%20Conduct-2.1-4baaaa)](https://github.com/ivancarlosti/awssesconverter?tab=coc-ov-file)
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors&color=ffc0cb)][sponsor]
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00)][buymeacoffee]
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

[cc]: https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/adding-a-code-of-conduct-to-your-project
[contributing]: https://docs.github.com/en/articles/setting-guidelines-for-repository-contributors
[security]: https://docs.github.com/en/code-security/getting-started/adding-a-security-policy-to-your-repository
[support]: https://docs.github.com/en/articles/adding-support-resources-to-your-project
[it]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/configuring-issue-templates-for-your-repository#configuring-the-template-chooser
[prt]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/creating-a-pull-request-template-for-your-repository
[funding]: https://docs.github.com/en/articles/displaying-a-sponsor-button-in-your-repository
[ivancarlos]: https://ivancarlos.me
[buymeacoffee]: https://buymeacoffee.com/ivancarlos
[paypal]: https://icc.gg/donate
[sponsor]: https://github.com/sponsors/ivancarlosti

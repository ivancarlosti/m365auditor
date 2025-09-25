# Microsoft 365 Auditor script
This script collects users, groups and Teams of a Microsoft 365 environment on .xlsx file for audit and review purposes

<!-- buttons -->
[![Stars](https://img.shields.io/github/stars/ivancarlosti/m365auditor?label=⭐%20Stars&color=gold&style=flat)](https://github.com/ivancarlosti/m365auditor/stargazers)
[![Watchers](https://img.shields.io/github/watchers/ivancarlosti/m365auditor?label=Watchers&style=flat&color=red)](https://github.com/sponsors/ivancarlosti)
[![Forks](https://img.shields.io/github/forks/ivancarlosti/m365auditor?label=Forks&style=flat&color=ff69b4)](https://github.com/sponsors/ivancarlosti)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/ivancarlosti/m365auditor?label=Activity)](https://github.com/ivancarlosti/m365auditor/pulse)
[![GitHub Issues](https://img.shields.io/github/issues/ivancarlosti/m365auditor?label=Issues&color=orange)](https://github.com/ivancarlosti/m365auditor/issues)
[![License](https://img.shields.io/github/license/ivancarlosti/m365auditor?label=License)](LICENSE)  
[![GitHub last commit](https://img.shields.io/github/last-commit/ivancarlosti/m365auditor?label=Last%20Commit)](https://github.com/ivancarlosti/m365auditor/commits)
[![Security](https://img.shields.io/badge/Security-View%20Here-purple)](https://github.com/ivancarlosti/m365auditor/security)
[![Code of Conduct](https://img.shields.io/badge/Code%20of%20Conduct-2.1-4baaaa)](https://github.com/ivancarlosti/m365auditor?tab=coc-ov-file)
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors&color=ffc0cb)][sponsor]
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
* For commercial related questions, please contact me directly for consulting costs. 

## 🩷 Project support
| If you found this project helpful, consider |
| :---: |
[**buying me a coffee**][buymeacoffee], [**donate by paypal**][paypal], [**sponsor this project**][sponsor] or just [**leave a star**](../..)⭐
|Thanks for your support, it is much appreciated!|

## 🌐 Connect with me
[![LinkedIn](https://img.shields.io/badge/LinkedIn-@ivancarlos-0077B5)](https://www.linkedin.com/in/ivancarlos)
[![X](https://img.shields.io/badge/X-@ivancarlos-000000)](https://x.com/ivancarlos)  
[![Signal](https://img.shields.io/badge/Signal-@ivancarlos.01-2592E9)](https://icc.gg/signal)
[![Telegram](https://img.shields.io/badge/Telegram-@ivancarlos-26A5E4)](https://t.me/ivancarlos)  
[![Discord](https://img.shields.io/badge/Discord-@ivancarlos.me-5865F2)](https://icc.gg/discord)
[![Website](https://img.shields.io/badge/Website-ivancarlos.me-FF6B6B)](https://ivancarlos.me)

[cc]: https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/adding-a-code-of-conduct-to-your-project
[contributing]: https://docs.github.com/en/articles/setting-guidelines-for-repository-contributors
[security]: https://docs.github.com/en/code-security/getting-started/adding-a-security-policy-to-your-repository
[support]: https://docs.github.com/en/articles/adding-support-resources-to-your-project
[it]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/configuring-issue-templates-for-your-repository#configuring-the-template-chooser
[prt]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/creating-a-pull-request-template-for-your-repository
[funding]: https://docs.github.com/en/articles/displaying-a-sponsor-button-in-your-repository
[ivancarlos]: https://ivancarlos.me
[buymeacoffee]: https://www.buymeacoffee.com/ivancarlos
[paypal]: https://icc.gg/donate
[sponsor]: https://github.com/sponsors/ivancarlosti

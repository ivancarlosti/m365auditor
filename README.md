# Microsoft 365 Auditor script
This script collects users, groups, Teams of a [Microsoft 365 Business]([https://workspace.google.com/](https://www.microsoft.com/microsoft-365/business)) environment on .xlsx file for audit and review purposes, the file is archived in a .zip file including a screenshot with hash MD5 of the .xlsx file and the script executed.

[![Stars](https://img.shields.io/github/stars/ivancarlosti/m365auditor?label=⭐%20Stars&color=gold&style=flat)](https://github.com/ivancarlosti/m365auditor/stargazers)
[![GitHub last commit](https://img.shields.io/github/last-commit/ivancarlosti/m365auditor?label=Last%20Commit)](https://github.com/ivancarlosti/m365auditor/commits)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/ivancarlosti/m365auditor?label=Activity)](https://github.com/ivancarlosti/m365auditor/pulse)
[![GitHub Issues](https://img.shields.io/github/issues/ivancarlosti/m365auditor?label=Issues&color=orange)](https://github.com/ivancarlosti/m365auditor/issues)  
[![License](https://img.shields.io/github/license/ivancarlosti/m365auditor?label=License)](LICENSE)
[![Security](https://img.shields.io/badge/Security-View%20Here-purple)](https://github.com/ivancarlosti/m365auditor/security)
[![Code of Conduct](https://img.shields.io/badge/Code%20of%20Conduct-1.4-4baaaa)](https://github.com/ivancarlosti/m365auditor/tree/main?tab=coc-ov-file)

## Instructions
* Save .ps1, .txt files locally (download [here](https://github.com/ivancarlosti/m365auditor/zipball/master))
* Change variables of `mainscript.ps1` if needed
* Run `mainscript.ps1` on PowerShell (right-click on file > Run with PowerShell)
* Follow instructions selecting project name, option 1 to generate audit report and collect .zip file on `$destinationpath`

## Requirements
* Windows 10+ or Windows Server 2019+
* PowerShell
* Modules `AzureAD`, `MicrosoftTeams`, `ImportExcel` on PowerShell

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
[![Discord](https://img.shields.io/badge/Discord-@ivancarlos.me-5865F2?logo=discord&logoColor=white)](https://discord.com/users/ivancarlos.me)
[![Signal](https://img.shields.io/badge/Signal-@ivancarlos.01-2592E9?logo=signal&logoColor=white)](https://icc.gg/-signal)
[![Telegram](https://img.shields.io/badge/Telegram-@ivancarlos-26A5E4?logo=telegram&logoColor=white)](https://t.me/ivancarlos)  
[![Instagram](https://img.shields.io/badge/Instagram-@ivancarlos-E4405F?logo=instagram&logoColor=white)](https://instagram.com/ivancarlos)
[![Threads](https://img.shields.io/badge/Threads-@ivancarlos-000000?logo=threads&logoColor=white)](https://threads.net/@ivancarlos)
[![X](https://img.shields.io/badge/X-@ivancarlos-000000?logo=x&logoColor=white)](https://x.com/ivancarlos)  
[![Website](https://img.shields.io/badge/Website-ivancarlos.me-FF6B6B?logo=linktree&logoColor=white)](https://ivancarlos.me)
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors&logo=githubsponsors&logoColor=white&color=ffc0cb)][sponsor]

## 📃 License
[MIT](LICENSE) © [Ivan Carlos][ivancarlos]

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

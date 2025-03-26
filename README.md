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

---

## Donation

| If you found this project helpful, consider |
| :---: |
[**buying me a coffee**][buymeacoffee], [**donate by paypal**][paypal], [**sponsor this project**][sponsor] or just [**leave a star**](../..)⭐
|Thanks for your support, it is much appreciated!|

## License

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

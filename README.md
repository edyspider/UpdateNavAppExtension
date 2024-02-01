# UpdateNavAppExtension

Safely Install and/or Upgrade NavApp Extension

## Prerequisites

* Powershell >= 2.0
* Microsoft Dynamics NAV >= 2018

This script uses the NAV Service CmdLets to execute the NavApp commands
* C:\Program Files\Microsoft Dynamics NAV\110\Service\NavAdminTool.ps1
* C:\Program Files\Microsoft Dynamics 365 Business Central\1X0\Service\NavAdminTool.ps1

---

## Script Validations

This script execute different tasks taking into consideration the following scenarios:
* If a newer version of the NavApp already exist
* If an equal version of the NavApp exist
* If older versions og the NavApp exists
* If has dependencies

---

## Example

```powershell
$navAppPath = 'C:\..\EdySpider_UpdateNavApp_23.0.0.0.app'
$navAppDeps = 'C:\..\Dependencies'
Update-NavApp -NavAppPath $navAppPath -NavService BC230 -NavVersion BC230 -SyncMode ForceSync -UnpublishOldVersion -NavAppDependenciesDir $navAppDeps.
```
Check the commented code at the end of the script
---

## Parameters

* Mandatory:

```powershell
-NavVersion             # NAV Version (NAV2018, BC130, BC140, BC150, BC160, BC170, BC180, BC190, BC200, BC210, BC220, BC230)
-NavService             # NAV Server Instance
-NavAppPath             # NavApp full file path
```

* Others Parameters

```powershell
-SyncMode               # NavApp Synchronize Mode
-UnpublishOldVersion    # Unpublish old versions
-NavAppDependenciesDir  # NavApp depencencies extensions directory (e.g.: Customer extension customizations)
```

---

## Authors

* [**EdySpider**](https://github.com/edyspider)

See also the list of [contributors](https://github.com/edyspider/UpdateNavAppExtension/contributors) who participated in this project.

---

## License

[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)

* **[MIT license](https://github.com/edyspider/UpdateNavAppExtension/blob/master/LICENSE)**
* Copyright 2021 © <a href="https://github.com/edyspider/" target="_blank">EdySpider
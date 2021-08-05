# UpdateNavAppExtension
Safetly Install or Upgrade NavApp Extension


### Prerequisites
* Powershell >= 2.0
* Microsoft Dynamics NAV >= 2018

This script uses the NAV Service CmdLets to execute the NavApp commands
 - C:\Program Files\Microsoft Dynamics NAV\110\Service\NavAdminTool.ps1
 - C:\Program Files\Microsoft Dynamics 365 Business Central\1X0\Service\NavAdminTool.ps1

---

## Configuration
Edit the parameters of the function ``Update-NavApp`` (last line) in the `Update-NavApp.ps1` file.

---

## Parameters
* Mandatory:
```powershell
-NavAppPath             # NavApp full file path
-NavService             # NAV Server Instance
-NavVersion             # NAV Version (NAV2018, BC130, BC140, BC150, BC160, BC170, BC180)
```

* Others Paramenters
```powershell
-SyncMode               # NavApp Synchronize Mode
-UnpublishOldVersion    # Unpublish old versions
```

---

## Authors

* [**EdySpider**](https://github.com/edyspider)

See also the list of [contributors](https://github.com/edyspider/UpdateNavAppExtension/contributors) who participated in this project.

---

## License

[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)

- **[MIT license](https://github.com/edyspider/UpdateNavAppExtension/blob/master/LICENSE)**
- Copyright 2021 © <a href="https://github.com/edyspider/" target="_blank">EdySpider

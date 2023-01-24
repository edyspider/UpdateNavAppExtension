<# ============================================================
 -- Author:       Éder Leal da Silva
 -- Create Date:  21/07/2021
 -- Description:  Safetly Install/Update NAV App Extension

 **************************
 ** Change History
 **************************
 ** PR  Date	     Author     Description	
 ** --  ----------  ----------  -------------------------------
 ** 01   21/07/2021  ENS        Create funcitons
 ** 02   05/08/2021  ENS        Add Synchronize Mode as Param
============================================================ #>

# Change module path
Import-Module 'C:\NAVHR\AppScripts\UpdateNavApp.ps1'


# Call Update-NavApp Script
$navAppPath = 'C:\Temp\AppSource\Arquiconsult_NAVHR_17.0.0.49.app'
$navAppDeps = 'C:\NAVHR\AppSource\Dependencies'
Update-NavApp -NavAppPath $navAppPath -NavService BC210 -NavVersion BC210 -SyncMode ForceSync -UnpublishOldVersion -NavAppDependenciesDir $navAppDeps


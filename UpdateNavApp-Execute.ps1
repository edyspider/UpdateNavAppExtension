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
 ** 03   07/02/2023  ENS        Change Code Example
 ** 04   07/01/2024  ENS        Change Code Example
============================================================ #>

# Change module path
Import-Module 'C:..\Scripts\PowerShell\UpdateNavApp\UpdateNavApp.ps1'


<#-------------------------------------------------------------------------------------------------------------------------------
 Use the following execution when using dependencies
-------------------------------------------------------------------------------------------------------------------------------#>
throw 'Please select the part of the script that you want to run' # Do not select this line

$navAppPath = 'W:\..\AL\AppSource\..\BusinessCentralExtension_23.0.0.0.app'
$navAppDeps = 'W:\..\AL\AppSource\..\Dependencies'

Update-NavApp -NavAppPath $navAppPath `
    -NavService DEMO230DEV -NavVersion BC230 -SyncMode ForceSync `
    -UnpublishOldVersion -NavAppDependenciesDir $navAppDeps -Verbose
<#-----------------------------------------------------------------------------------------------------------------------------#>



<#-------------------------------------------------------------------------------------------------------------------------------
 Use the following execution when not using dependencies
-------------------------------------------------------------------------------------------------------------------------------#>
throw 'Please select the part of the script that you want to run' # Do not select this line

$navAppPath = 'W:\..\AL\AppSource\..\BusinessCentralExtension_23.0.0.0.app'

Update-NavApp -NavAppPath $navAppPath `
    -NavService DEMO230DEV -NavVersion BC230 -SyncMode ForceSync -UnpublishOldVersion -Verbose
<#-----------------------------------------------------------------------------------------------------------------------------#>
<# ============================================================
 -- Author:       EdySpider
 -- Create Date:  21/07/2021
 -- Description:  Safetly Install/Update NAV App Extension

 **************************
 ** Change History
 **************************
 ** PR  Date	     Author     Description	
 ** --  ----------  ----------  -------------------------------
 ** 01   21/07/2021  ENS        Create funcitons
============================================================ #>

function Update-NavApp () {
    param(
        [parameter(Mandatory=$true)]
        [ValidateSet("NAV2018","BC130","BC140","BC150","BC160","BC170","BC180")]
        [string]$NavVersion,

        [parameter(Mandatory=$true)]
        [string]$NavService,

        [parameter(Mandatory=$true)]
        [string]$NavAppPath,

        [parameter(Mandatory=$false)]
        [switch]$UnpublishOldVersion
    )

    $statusActivity = "Updating NavApp Extension";

    $status = "Loading Nav Modules.."
    Show-ProgressBar -Activity $statusActivity -Status $status -Position 10
    Start-Sleep -Seconds 2

    # Load Nav Modules
    switch ($NavVersion) {
        NAV2018 {
            Import-Module 'C:\Program Files\Microsoft Dynamics NAV\110\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC130 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\130\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC140 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\140\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC150 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\150\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC160 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\160\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC170 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\170\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
        BC180 {
            Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\180\Service\NavAdminTool.ps1' -WarningAction SilentlyContinue | Out-Null
        }
    }

    # Get NavApp Information
    $status = "Getting NavApp Information..."
    Show-ProgressBar -Activity $statusActivity -Status $status -Position 25
    Start-Sleep -Seconds 2

    $appInfo = (Get-NAVAppInfo -Path $NavAppPath | Select-Object)
    $appInfo | Select-Object -Property Name, Version, Publisher

    $appId = $appInfo.AppId.ToString()
    $appName = $appInfo.Name.ToString()
    $appVersion = $appInfo.Version.ToString()
    $appPublisher = $appInfo.Publisher.ToString()

    # Check If NavApp Already Exists
    $status = "Checking Existing NavApp.."
    Show-ProgressBar -Activity $statusActivity -Status $status -Position 30
    Start-Sleep -Seconds 2
    $existApp = (Get-NAVAppInfo -ServerInstance $NavService -Id $appId) | Sort-Object -Property Version | Select-Object -Last 1

    if ($existApp) {
        $existAppName = $existApp.Name
        $exAppVersion = $existApp.Version.ToString()

        # Uninstall Existing NavApp
        $status = "Uninstall Existing NavApp..."
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 40
        Start-Sleep -Seconds 2
        UnInstall-NAVApp -ServerInstance $NavService -Name $existAppName -version $exAppVersion
    
        if($appVersion -eq $exAppVersion) {
            # Unpublish Existing NavApp
            $status = "Unpublishing Old NavApp Version..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 50
            Start-Sleep -Seconds 2
            Unpublish-NAVApp -ServerInstance $NavService -Name $existAppName -Version $exAppVersion
            
            # Publish New NavApp
            $status = "Publishing New NavApp Version..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 65
            Start-Sleep -Seconds 2
            Publish-NAVApp -ServerInstance $NavService -Path $NavAppPath -PassThru -SkipVerification

            # Sync-NavApp
            $status = "Synchronizing NavApp..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 80
            Start-Sleep -Seconds 2
            Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion

            # Install NavApp
            $status = "Installing NavApp..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 90
            Start-Sleep -Seconds 2
            Install-NAVApp -ServerInstance $NavService -Name $appName -version $appVersion

        } else {
            # Publish New NavApp
            $status = "Publishing New NavApp Version..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 60
            Start-Sleep -Seconds 2
            Publish-NAVApp -ServerInstance $NavService -Path $NavAppPath -PassThru -SkipVerification
          
            # Sync-NavApp
            $status = "Synchronizing NavApp..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 75
            Start-Sleep -Seconds 2
            Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion
          
            # Upgrade NavApp
            $status = "Upgrading NavApp..."
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 85
            Start-Sleep -Seconds 2
            Start-NAVAppDataUpgrade -ServerInstance $NavService -Name $appName -version $appVersion

            # Unpublish NavApp
            if($UnpublishOldVersion) {
                $status = "Unpublishing Old NavApp Version..."
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 90
                Start-Sleep -Seconds 2
                Unpublish-NAVApp -ServerInstance $NavService -Name $appName -Version $exAppVersion
            }
        }
    } else {
        # Publish New NavApp
        $status = "Publishing NavApp..."
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 50
        Start-Sleep -Seconds 2
        Publish-NAVApp -ServerInstance $NavService -Path $NavAppPath -PassThru -SkipVerification

        # Sync-NavApp
        $status = "Synchronizing NavApp..."
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 70
        Start-Sleep -Seconds 2
        Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion
        
        # Install NavApp
        $status = "Installing NavApp..."
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 90
        Start-Sleep -Seconds 2
        Install-NAVApp -ServerInstance $NavService -Name $appName -Version $appVersion #-Publisher $appPublisher
    }

    $status = "Finishing Process..."
    Show-ProgressBar -Activity $statusActivity -Status $status -Position 97
    Start-Sleep -Seconds 2
}

function Show-ProgressBar{
    param (
        [parameter(Mandatory=$true)]
        [string]$Activity,
        [parameter(Mandatory=$false)]
        [string]$Status,
        [parameter(Mandatory=$true)]
        [string]$Position
    )

    Write-Progress -Activity $Activity -Status "$Status - $Position% Complete:" -PercentComplete $Position;
}

Update-NavApp -NavAppPath 'C:\..\EdySpider_UpdateNavApp_1.0.0.0.app' -NavService BC170 -NavVersion BC170 -UnpublishOldVersion

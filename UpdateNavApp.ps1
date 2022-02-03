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
 ** 02   05/08/2021  ENS        Add Synchronize Mode as Param
 ** 03   21/09/2021  ENS        Remove Multiple Old Versions
 ** 04   21/09/2021  ENS        Republish NavApp Dependencies
 ** 05   03/02/2022  ENS        Include BC 19.0 version
============================================================ #>

Write-Host "Welcome to the UpdateNavApp Script by EdySpider" -ForegroundColor Green
Write-Host "For more information please go to 'https://github.com/edyspider/UpdateNavAppExtension'" -ForegroundColor DarkCyan
Write-Host ' '

function Update-NavApp () {
    param(
        [parameter(Mandatory=$true)]
        [ValidateSet("NAV2018","BC130","BC140","BC150","BC160","BC170","BC180", "BC190")]
        [string]$NavVersion,

        [parameter(Mandatory=$true)]
        [string]$NavService,

        [parameter(Mandatory=$true)]
        [string]$NavAppPath,

        [parameter(Mandatory=$false)]
        [ValidateSet("Add","Clean","Devlopment","ForceSync")]
        [string]$SyncMode,

        [parameter(Mandatory=$false)]
        [switch]$UnpublishOldVersion,

        [parameter(Mandatory=$false)]
        [string]$NavAppDependenciesDir
    )

    Try
    {
        $statusActivity = "Updating NavApp Extension";
        $status = "Loading Nav Modules.."
        Write-Host $status -ForegroundColor Yellow
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 10
        Start-Sleep -Seconds 1

        # Load Nav Modules
        switch ($NavVersion) {
            NAV2018 {
                Import-Module 'C:\Program Files\Microsoft Dynamics NAV\110\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynNAV 2018 Module" -ForegroundColor White
            }
            BC130 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\130\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 13.0 Module" -ForegroundColor White
            }
            BC140 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\140\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 14.0 Module" -ForegroundColor White
            }
            BC150 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\150\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 15.0 Module" -ForegroundColor White
            }
            BC160 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\160\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 16.0 Module" -ForegroundColor White
            }
            BC170 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\170\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 17.0 Module" -ForegroundColor White
            }
            BC180 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\180\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 18.0 Module" -ForegroundColor White
            }
            BC190 {
                Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\190\Service\NavAdminTool.ps1' -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host  "Load DynBC 365 19.0 Module" -ForegroundColor White
            } else {
                throw 'NavVersion parameter is not valid for Update-NavApp'
            }
        }

        Write-Host ' '
        # Get NavApp Information
        $status = "Getting NavApp Information..."
        Write-Host $status -ForegroundColor Yellow
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 25
        Start-Sleep -Seconds 1

        $appInfo = (Get-NAVAppInfo -Path $NavAppPath | Select-Object)

        $appId = $appInfo.AppId.ToString()
        $appName = $appInfo.Name.ToString()
        $appVersion = $appInfo.Version.ToString()
        $appPublisher = $appInfo.Publisher.ToString()

        $global:upgradeApp = 1

        Write-Host 'Name:     ' $appName -ForegroundColor White
        Write-Host 'Version:  ' $appVersion -ForegroundColor White
        Write-Host 'Publisher:' $appPublisher -ForegroundColor White

        Write-Host ' '        
        # Check If NavApp Already Exists
        $status = "Checking Existing " + $appName + " NavApp..."
        Write-Host $status -ForegroundColor Yellow
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 30
        Start-Sleep -Seconds 1
    
        $lastExistApp = (Get-NAVAppInfo -ServerInstance $NavService -Id $appId) | Sort-Object -Property Version | Select-Object -Last 1

        if ($lastExistApp) {
            $lastAppVersion = $lastExistApp.Version.ToString()
            if ($lastAppVersion -gt $appVersion) {
                throw 'This database already has a more recent version of ' + $appName + ' ('+$lastAppVersion+'). If the current NavApp if not working, please uninstall and unpublish manually.'
            }

            Get-ChildItem -Path $NavAppDependenciesDir -Filter *.app -Recurse -File -Name| ForEach-Object {
                #$NavAppDependency = [System.IO.Path]::GetFullPath($_)
                $NavAppDependency = $NavAppDependenciesDir + '\' + $_
                $DepAppInfo = (Get-NAVAppInfo -Path $NavAppDependency | Select-Object)

                $DepAppId = $DepAppInfo.AppId.ToString()
                $DepAppName = $DepAppInfo.Name.ToString()
                $DepAppVersion = $DepAppInfo.Version.ToString()

                # Uninstall Dedendent NavApp
                $status = "Uninstall Dedendent NavApp..."
                Write-Host $status -ForegroundColor Yellow
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 40
                Start-Sleep -Seconds 1
                UnInstall-NAVApp -ServerInstance $NavService -Name $DepAppName -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host $DepAppName " uninstalled!" -ForegroundColor White

                # Unpublish Dedendent NavApp
                $status = "Unpublishing Dedendent NavApp..."
                Write-Host $status -ForegroundColor Yellow
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 40
                Start-Sleep -Seconds 1
                Unpublish-NAVApp -ServerInstance $NavService -Name $DepAppName -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host $DepAppName " unpublished!" -ForegroundColor White
            }
        
            $OldVersions = Get-NAVAppInfo $NavService -Name $appName | Sort-Object -Property Version
            $OldVersions | ForEach-Object  {
                $existAppName = $_.Name
                $exAppVersion = $_.Version.ToString()

                # Uninstall Existing NavApp
                $status = "Uninstall Existing NavApp..."
                Write-Host $status -ForegroundColor Yellow
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 45
                Start-Sleep -Seconds 1
                UnInstall-NAVApp -ServerInstance $NavService -Name $existAppName -version $exAppVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host $existAppName $exAppVersion ' uninstalled!' -ForegroundColor White

                if($appVersion -eq $exAppVersion) {
                    $upgradeApp = 0

                    # Unpublish Existing NavApp
                    $status = "Unpublishing NavApp with the same Version..."
                    Write-Host $status -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                    Unpublish-NAVApp -ServerInstance $NavService -Name $existAppName -Version $exAppVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                    Write-Host $existAppName $exAppVersion ' unpublished!' -ForegroundColor White
                }
            }

            Write-Host ' '
            # Publish New NavApp
            $status = "Publishing NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 50
            Start-Sleep -Seconds 1
            Publish-NAVApp -ServerInstance $NavService -Path $NavAppPath -PassThru -SkipVerification -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Host $appName ' published!' -ForegroundColor White

            Write-Host ' '
            # Sync-NavApp
            $status = "Synchronizing NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 70
            Start-Sleep -Seconds 1
            SyncNavApp -NavService $NavService -AppName $appName -AppVersion $appVersion -SyncMode $SyncMode
            Write-Host $appName " synchronized!" -ForegroundColor White
        
            Write-Host ' '
            if ($upgradeApp -eq 1) {
                # Upgrade NavApp
                $status = "Upgrading NavApp..."
                Write-Host $status -ForegroundColor Yellow
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 85
                Start-Sleep -Seconds 2
                Start-NAVAppDataUpgrade -ServerInstance $NavService -Name $appName -version $appVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host $appName " upgraded!" -ForegroundColor White
            } else {
                # Install NavApp
                $status = "Installing NavApp..."
                Write-Host $status -ForegroundColor Yellow
                Show-ProgressBar -Activity $statusActivity -Status $status -Position 85
                Start-Sleep -Seconds 1
                Install-NAVApp -ServerInstance $NavService -Name $appName -Version $appVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                Write-Host $appName " installed!" -ForegroundColor White
            }

            # Unpublish NavApp
            if($UnpublishOldVersion) {
                $OldVersions = Get-NAVAppInfo $NavService -Name $appName | Sort-Object -Property Version
                $OldVersions | ForEach-Object  {

                    $existAppName = $_.Name
                    $exAppVersion = $_.Version.ToString()

                    if ($exAppVersion -lt $appVersion) {
                        $status = "Unpublishing Old NavApp Version..."
                        Write-Host $status -ForegroundColor Yellow
                        Show-ProgressBar -Activity $statusActivity -Status $status -Position 90
                        Start-Sleep -Seconds 2
                        Unpublish-NAVApp -ServerInstance $NavService -Name $appName -Version $exAppVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                        Write-Host $appName $exAppVersion " unpublished!" -ForegroundColor White
                    }
                }
            }

        } else {
            Write-Host ' '
            # Publish New NavApp
            $status = "Publishing NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 50
            Start-Sleep -Seconds 1
            Publish-NAVApp -ServerInstance $NavService -Path $NavAppPath -PassThru -SkipVerification -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Host $appName " published!" -ForegroundColor White

            Write-Host ' '
            # Sync-NavApp
            $status = "Synchronizing NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 70
            Start-Sleep -Seconds 1
            SyncNavApp -NavService $NavService -AppName $appName -AppVersion $appVersion -SyncMode $SyncMode
            Write-Host $appName " synchronized!" -ForegroundColor White
        
            Write-Host ' '
            # Install NavApp
            $status = "Installing NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 90
            Start-Sleep -Seconds 1
            Install-NAVApp -ServerInstance $NavService -Name $appName -Version $appVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null #-Publisher $appPublisher
            Write-Host $appName " installed!" -ForegroundColor White
        }

        Get-ChildItem -Path $NavAppDependenciesDir -Filter *.app -Recurse -File -Name| ForEach-Object {
            #$NavAppDependency = [System.IO.Path]::GetFullPath($_)
            $NavAppDependency = $NavAppDependenciesDir + '\' + $_
            $DepAppInfo = (Get-NAVAppInfo -Path $NavAppDependency | Select-Object)

            $DepAppId = $DepAppInfo.AppId.ToString()
            $DepAppName = $DepAppInfo.Name.ToString()
            $DepAppVersion = $DepAppInfo.Version.ToString()

            Write-Host ' '
            # Publish New NavApp
            $status = "Publishing Dependent NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 95
            Start-Sleep -Seconds 1
            Publish-NAVApp -ServerInstance $NavService -Path $NavAppDependency -PassThru -SkipVerification -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Host $DepAppName " published!" -ForegroundColor White

            Write-Host ' '
            # Sync-NavApp
            $status = "Synchronizing Dependent NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 95
            Start-Sleep -Seconds 1
            SyncNavApp -NavService $NavService -AppName $DepAppName -AppVersion $DepAppVersion -SyncMode $SyncMode
            Write-Host $DepAppName " synchronized!" -ForegroundColor White
        
            Write-Host ' '
            # Install NavApp
            $status = "Installing Dependent NavApp..."
            Write-Host $status -ForegroundColor Yellow
            Show-ProgressBar -Activity $statusActivity -Status $status -Position 95
            Start-Sleep -Seconds 1
            Install-NAVApp -ServerInstance $NavService -Name $DepAppName -Version $DepAppVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null #-Publisher $appPublisher
            Write-Host $DepAppName " installed!" -ForegroundColor White
        }

        Write-Host ' '
        $status = "Finishing Process..."
        Write-Host $status -ForegroundColor Yellow
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 97
        Start-Sleep -Seconds 1
        Show-ProgressBar -Activity $statusActivity -Status $status -Position 100
        Write-Host 'All Done!' -ForegroundColor White
    }
    Catch
    {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName

        Write-Host 'UpdateNavApp Failed!' -ForegroundColor Red -BackgroundColor Black
        Write-Host "We failed to read file $FailedItem. The error message was $ErrorMessage" -ForegroundColor Red -BackgroundColor Black
        Break
    }
}

function SyncNavApp () {
    param(
        [parameter(Mandatory=$true)]
        [string]$NavService,

        [parameter(Mandatory=$true)]
        [string]$AppName,

        [parameter(Mandatory=$true)]
        [string]$AppVersion,

        [parameter(Mandatory=$false)]
        [ValidateSet("Add","Clean","Devlopment","ForceSync")]
        [string]$SyncMode
    )

    if($SyncMode) {
        switch ($SyncMode) {
            Add {
                Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion -Mode Add -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Force | Out-Null
            }
            Clean {
                Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion -Mode Clean -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Force | Out-Null
            }
            Devlopment {
                Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion -Mode Development -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Force | Out-Null
            }
            ForceSync {
                Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion -Mode ForceSync -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Force | Out-Null
            }
        }
    } else {
        Sync-NavApp -ServerInstance $NavService -Name $appName -version $appVersion -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Force | Out-Null
    }
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

# Call Update-NavApp Script
#$navAppPath = 'C:\..\EdySpider_UpdateNavApp_1.0.0.0.app'
#$navAppDeps = 'C:\..\Dependencies'
#Update-NavApp -NavAppPath $navAppPath -NavService BC170 -NavVersion BC170 -SyncMode ForceSync -UnpublishOldVersion -NavAppDependenciesDir $navAppDeps

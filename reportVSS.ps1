#Requires -RunAsAdministrator
$ErrorActionPreference = "Stop"


Function funVssSanityCheck
    {

        $Volumes = Get-WmiObject -Class "Win32_Volume" -Property * -Filter "DriveType=3"
        $ShadowStorage = Get-WmiObject -Class "Win32_ShadowStorage" -Property *
        $ShadowCopies = Get-WmiObject -Class "Win32_ShadowCopy" -Property *


        # convert something like 20170822115910.108925+600 to DateTime
        Function funInstallDate2DateTime
            {
                Param([Parameter(Mandatory=$true, ValueFromPipeline=$true)][String]$String)
                Return ([datetime]::ParseExact($String.Split('.')[0], 'yyyyMMddHHmmss', $null))
            }


        Function funDaysOld
            {
                Param([Parameter(Mandatory=$true, ValueFromPipeline=$true)][DateTime]$DateTime)
                Return (New-TimeSpan -Start ($DateTime) -End (Get-Date) | Select-Object -ExpandProperty "Days")
            }


        # expected input, something like:     Win32_Volume.DeviceID="\\\\?\\Volume{8dfba37a-95db-11e4-80bf-001dd8b71c0b}\\"     -or-     \\?\Volume{8dfba37a-95db-11e4-80bf-001dd8b71c0b}\
        # expected output, something like:    8dfba37a-95db-11e4-80bf-001dd8b71c0b
        Function funRegExDeviceID
            {
                Param([Parameter(Mandatory=$true, ValueFromPipeline=$true)][String]$TheThing)
                $TheThing -match "{(.*)}" | Out-Null

                Return ($Matches[1])
            }


        $ResultsTable = New-Object System.Data.DataTable
        foreach ($i_Volume in ($Volumes | Where-Object -FilterScript {$PSItem.DriveLetter -ne $null}))
            {
                $zzVolumeLetter = $i_Volume.DriveLetter
                $zzVolumeLabel = $i_Volume.Label
                $zzDeviceID = $i_Volume.DeviceID
                $zzNumberOfShadows = $ShadowCopies | Where-Object -FilterScript {$PSItem.VolumeName -eq $i_Volume.DeviceID} | Measure-Object | Select-Object -ExpandProperty "Count"
        
                Switch ($zzNumberOfShadows -eq 0)
                    {
                        $true
                            {
                                $zzOldestShadow = $null
                                $zzNewestShadow = $null
                                $zzNewestShadowDaysOld = $null
                                $zzShadowLocation = $null
                            }

                        $false
                            {
                                $zzOldestShadow = $ShadowCopies | Where-Object -FilterScript {$PSItem.VolumeName -eq $i_Volume.DeviceID } | Sort-Object -Property "InstallDate" | Select-Object -First 1 | Select-Object -ExpandProperty "InstallDate" | funInstallDate2DateTime
                                $zzNewestShadow = $ShadowCopies | Where-Object -FilterScript {$PSItem.VolumeName -eq $i_Volume.DeviceID } | Sort-Object -Property "InstallDate" | Select-Object -Last 1 | Select-Object -ExpandProperty "InstallDate" | funInstallDate2DateTime
                                $zzNewestShadowDaysOld = $zzNewestShadow | funDaysOld
                        
                                $zzDiffVolumeDeviceID = $ShadowStorage | Where-Object -FilterScript {$PSItem.Volume -like "*$(funRegExDeviceID -TheThing $i_Volume.DeviceID)*"} | Select-Object -ExpandProperty "DiffVolume"
                                $zzShadowLocationLetter = $Volumes | Where-Object -FilterScript {$PSItem.DeviceID -like "*$((funRegExDeviceID -TheThing $zzDiffVolumeDeviceID))*"} | Select-Object -ExpandProperty "DriveLetter"
                                $zzShadowLocationLabel = $Volumes | Where-Object -FilterScript {$PSItem.DriveLetter -eq $zzShadowLocationLetter} | Select-Object -ExpandProperty "Label"
                            }
                    }
        
        
                $ResultsTable += New-Object psobject -Property @{
                "VolumeLetter" = $zzVolumeLetter;
                "VolumeLabel" = $zzVolumeLabel;
                "NumberOfShadows" = $zzNumberOfShadows;
                "OldestShadow" = $zzOldestShadow
                "NewestShadow" = $zzNewestShadow;
                "NewestShadowDaysOld" = $zzNewestShadowDaysOld;
                "ShadowLocationLetter" = $zzShadowLocationLetter;
                "ShadowLocationLabel" = $zzShadowLocationLabel
                }
            }

        Write-Host $env:COMPUTERNAME
        Return $ResultsTable
    }

---
layout: default
title: System
parent: Documentation
nav_order: 2
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## GET-WMIOBJECT : Invalid class "WIN32_VOLUME"

```
GET-WMIOBJECT : Invalid class "WIN32_VOLUME"
    + CategoryInfo          : InvalidType: (:) [Get-WmiObject], ManagementException
    + FullyQualifiedErrorId : GetWMIManagementException,Microsoft.PowerShell.Commands.GetWmiObjectCommand
```

This error was encountered on a device that was pending upgrade to Windows 10 v1909, the validation task running to detect if the machine was eligible errored out when running this command and that made it seem like the machine had no disk space available. The WMI had to be completely reset to resolve the issue it had.

```
Stop-Service -Name "Winmgmt" -Force -Confirm:$false
cmd /c "winmgmt /resetrepository"
```

*Rebuilding the WMI repository may result in some third-party products not working until their setup is re-run & their Microsoft Operations Framework re-added back to the repository.*

## "The installer has encountered an unexpected error installing this package. This may indicate a problem with this package. The error code is 2755."

During installations all the files get extracted in to "C:\Windows\Installer", in this case when running installations nothing appeared to be happening when running the installation script. I just happened to run a .msi manually an received this error code, in this scenario the "Installer" folder was missing and was resolved by creating this folder in "C:\Windows".

```
Missing: "C:\Windows\Installer"
```

## Windows 10 v1909: Full HDD - C:\Windows\Temp\Microsoft-Windows\*.evtx

* C:\Windows\Temp - completely filled (330GB) with .evtx files
* Could be related to v1903 update KB4505903

```
C:\Windows\Temp\
Microsoft-Windows-AppReadiness_Admin_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-AppReadiness_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-AppXDeploymentServer_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-AppXPackaging_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-SettingSync_Debug_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-SettingSync_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-StateRepository_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-Store_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
Microsoft-Windows-WindowsUpdateClient_Operational_D5CD4A48-AC38-0000-44CE-D1D538ACD601.evtx
```

The below changes stopped further events from being generated.

```
Local GPO: User Configuration > Administrative Templates > Windows Components > Store > Enable > Turn off the Store application
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\AppXSvc' /v Start /t REG_DWORD /d 4 /f

reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' /v LetAppsGetDiagnosticInfo /t REG_DWORD /d 2 /f
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' /v LetAppsRunInBackground /t REG_DWORD /d 2 /f
```




















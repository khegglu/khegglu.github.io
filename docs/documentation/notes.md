---
layout: default
title: Notes
parent: Documentation
nav_order: 1
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## System:

```
# Get the system disk space
Get-WmiObject -Class Win32_logicaldisk -Filter "DriveType = '3'" | Select-Object -Property DeviceID, DriveType, VolumeName, @{L='FreeSpaceGB';E={"{0:N2}" -f ($_.FreeSpace /1GB)}}, @{L="Capacity";E={"{0:N2}" -f ($_.Size/1GB)}}


```



## Powershell:

```
# Clone group membership from 1 AD group to another
Add-ADGroupMember -Identity 'newgroup' -Members (Get-ADGroupMember -Identity 'oldgroup' -Recursive)

# Check if special password policy is applied on a account # If no result then the default domain policy applies: Get-ADDefaultDomainPasswordPolicy
Get-ADUserResultantPasswordPolicy $user


```

## Registry:

```
# Get Windows 10 OS version
(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId

# Disable flash EOL notification message in Internet Explorer, plugin will be removed at the end of next year.
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Main" /v DisableFlashNotificationPolicy /t REG_DWORD /d 1 /f
```

## Installs:

```
# Get SID's on computer
Reg Query 'HKEY_USERS'

# Get reg detals for apps installed in users profile
Reg Query "HKEY_USERS\S-1-5-21-X-X-X-X\software\microsoft\windows\currentversion\uninstall\"

```

```
# Silent install of MSODBCSQL
msiexec /i C:\Packages\msodbcsql_17.5.2.1_x64.msi /qn ADDLOCAL=ALL IACCEPTMSODBCSQLLICENSETERMS=YES

# Silent install of SQLNC
msiexec /i C:\Packages\sqlncli.msi /qn ADDLOCAL=ALL IACCEPTSQLNCLILICENSETERMS=YES

# Silent install of VNC Server
cmd /c 'c:\packages\VNC-Server-6.7.2-Windows.exe' /qn REBOOT=ReallySuppress LICENSEKEY=X-X-X-X-X ENABLEAUTOUPDATECHECKS=0 ENABLEANALYTICS=0
** License confirmation: cmd /c 'C:\Program Files\RealVNC\VNC Server\vnclicense.exe' -add X-X-X-X-X
# Silent install of VNC Viewer
cmd /c 'c:\packages\VNC-Viewer-6.20.529-Windows.exe' /qn

```



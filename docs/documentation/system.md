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

### You have been logged in with a temporary profile

```
# Remote connect to the device and remove all corrupted user profiles in the rgistry (key contains .bak)
Get-ChildItem -Path "hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Where Name -like "*.bak" | Remove-Item -confirm:$false
# Restart the machine, then clear out all "Temp" folders from C:\Users
cmd /c 'for /d %G in ("C:\Users\TEMP*") do rd /s /q "%~G"'
```

### Block driver updates - Windows 10 Audio Update issues

*For most devices there are more than one Hardware Ids. Usually the first two are the ones you need. The first one is the actual Hardware Id of your device and is specific to your device. For proper prevention, it is recommended to select the second one, which is more generic.*

```
# Locate Hardware ID:
devmgmt.msc > Device properties > Details tab > Hardware ID from the dropdown list
- PCI\VEN_8086&DEV_9D70&SUBSYS_8079103C&REV_21
- PCI\VEN_8086&DEV_9D70&SUBSYS_8079103C
* Select the second option, which is more generic rather than the specific device which is the first.

# Enable block list:
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions\DenyDeviceIDs\Restriction\DenyDeviceIDs" /v DenyDeviceIDs /t REG_DWORD /d 1 /f

# Set blacklist:
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions\DenyDeviceIDs\Restriction\DenyDeviceIDs" /v 1 /t REG_SZ /d "INTELAUDIO\FUNC_01&VEN_14F1&DEV_50F4&SUBSYS_103C8079" /f
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions\DenyDeviceIDs\Restriction\DenyDeviceIDs" /v 2 /t REG_SZ /d "PCI\VEN_8086&DEV_9D70&SUBSYS_8079103C" /f
```

## Bitlocker:
### Cannot retrieve recovery password information. Size limit for the request has been exceeded. (Active Directory)

This was a issue that presented itself on a newly reimaged device, the old object in AD was just re-established and it kept all the old childobjects. There must have been some issue with bitlocker at some point since the attribute on this machine displayed 1200+ recovery keys, issue was resolved by clearing these and syncing the new key.

Forums indicated that this issue is more likely to occur on machines where they use a lot of USB-sticks that is being encrypted.

```
# Computer with the issue
$computer = Get-ADComputer GBXXXWNX3X00000$
# Command to pull the existing list and print to screen
Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -SearchBase $computer.DistinguishedName -Properties whenCreated, msFVE-RecoveryPassword | Sort whenCreated -Descending | Select whenCreated, msFVE-RecoveryPassword
# Command to pull the existing list and remove the child objects
Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -SearchBase $computer.DistinguishedName -Properties whenCreated, msFVE-RecoveryPassword | Remove-ADObject -confirm:$false

# One liner to pull the existing data and print to screen
Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -SearchBase (Get-ADComputer GBSOTWDA3S05824$).DistinguishedName -Properties whenCreated, msFVE-RecoveryPassword | Sort whenCreated -Descending | Select whenCreated, msFVE-RecoveryPassword, @{Name="Name";Expression={(($_.Name).Split("{")[1]).split("-")[0]}}

# From the device force a resync to AD
manage-bde -protectors -get c:
manage-bde -protectors -adbackup c: -id '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'
# Alternative
$BLV = Get-BitLockerVolume -MountPoint "C:"
Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $BLV.KeyProtector[1].KeyProtectorId
```

### Failed to encrypt C: - Bitlocker could not encrypt one or more drives on this computer. You will be reminded to encrypt this computer again.

Generally this error is due to the Key protectors being missing, resolved by adding a key protector, then enabling bitlocker. The last feedback from the command should be that bitlocker will begin encrypting after doing a self hardware check.

```
# manage-bde -status
Add-BitLockerKeyProtector -MountPoint 'C:' -RecoveryPasswordProtector
manage-bde -on C: -RecoveryPassword
```

### Keeps asking for recovery key after every reboot

*After using a recovery key the protection needs to be suspended and resumed, otherwise it will ask for the recovery key on every reboot of the machine*

```
manage-bde.exe -protectors -disable c:
manage-bde.exe -protectors -enable c:
```

### The TPM is defending against dictionary attacks and is in a time-out period.

```
# To clear the TPM, when rebooting the machine the user needs to confirm on screen
Initialize-Tpm -AllowClear -AllowPhysicalPresence
```

### "Turn on Bitlocker" vanished for USB Drives

- [Reference](https://docs.microsoft.com/answers/answers/160728/view.html)

For Windows 10 v1909 the issue is caused by the patches: KB4577671/KB4586786 (removing the updates appears to work for now). An official fix has not been released as of yet!

*Windows Engineering has decided to revert this change after finding out that several manufacturers are shipping USB Sticks with partitions marked as Active.*

```
Workaround 1:
As already stated by JulianFloyd-5310, one solution is to delete the volume of the desired drive in Disk Managment and then creating a volume again. After that you need to plug the USB drive out and in again. The BitLocker option should show up in the context menu of the USB drive after that.

Workaround 2:
Another solution is done with diskpart. Use this solution if you want to keep the contents of your USB drive.
Simply open a new cmd, type in diskpart and confirm adminprompt. In diskpart type in "list disk" and locate your USB drive. Select your USB drive with "select disk ###" (replace "###" with the desired number). Then type in "list partition" and select the primary partition with "select partition ###". After selecting both disk and partition, type in "inactive". Your drive should be set to inactive now and after plugging the USB stick out and in again, the BitLocker option should be available again.
```

## System Devices:
### Intel Wireless Bluetooth - Currently, this hardware device is not connected to the computer. (Code 45) To fix this problem, reconnect this hardware device to the computer.

*The OS is not enumerating the device properly. This can be caused by physical trauma causing the device to become dislodged, or the device failing. More often than not, the hardware is just fine.*

This issue is usually resolved by a simple static drain. Power off the computer, then hold the power button for at least 1 minute, before you then power it back on again.

## The Group Policy Client service failed the logon. Access is denied.

Issue due to corrupt user registries for the affected user, proceed to remove the users profile in the registry and rename the users folder.

## Disable offline files

```
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\CscService" -Name "Start" -Value 1
Stop-Service CscService
Set-Service CscService -StartupType Disabled
reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Csc\Parameters /v FormatDatabase /t REG_DWORD /d 1 /f
```

## Unable to map shared drive in windows 10, offline files is off

```
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Mup\Parameters" /v EnableDfsLoopbackTargets /t REG_DWORD /d 1 /f
```

## Network drives disappearing
By default windows 10 will drop idle connections after a specified time-out period.
Open cmd as administrator and use the command below to turn the auto-disconnect feature off.
> net config server /autodisconnect:-1

```
# Registry alternative, can also query this key to verify current settings
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" /v autodisconnect /t REG_DWORD /d 0xf /f
```

## This share requires the obsolete SMB1 protocol

*As of a Windows 10 v1709 update the SMBv1 protocol is by default not installed, this is due to security and vulnerability issues. You will see this error if you are trying to connect to a share that is hosted on windows server 2003, or newer where they are not using SMBv2(2008) or SMBv3(2012).*

The error can be bypassed by enabling the protocol again.

```
Set-SmbServerConfiguration –EnableSMB1Protocol $true
```

## There is a problem with your Remote Desktop license, and your session will be disconnected in 60 minutes, Contact your system administrator

To resolve, remove the registry key that stores the RDS license. A new key should be generated once this has been removed.

```
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSLicensing" /f
```

## Windows 10 machine going to sleep exactly 2 mins after the screen has been locked

After some research the registry power option "System Unattended Sleep Timout" was discovered, this setting by default is set to 2 minutes matching the problem description. It's an issue that has affected only a few machines, so not sure what changed to start enforcing that policy.

```
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0" /v Attributes /t REG_DWORD /d 2 /f
```

After adding the key, open advance power settings and set the option: "System Unattended Sleep Timout" from it's 2 minutes to 0 to disable the feature.

```
Root: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0\DefaultPowerSchemeValues

In that location there are 3 system power options
381b4222-f694-41f0-9685-ff5bb260df2e
8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
a1841308-3541-4fab-bc81-f71556f20b4a

Containing the keys:
AcSettingIndex REG_DWORD 0x78
DcSettingIndex REG_DWORD 0x78

That 0x78 hex translates to 120 which is milliseconds so the 2 minute default timer. Making that change I did on the the users machine did not adjust those keys.

There is then a "User" section in HKLM you can pick from in the power options so I did a:
reg query 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\Powerschemes'

On mine and and the users machine it shows that we are both using the same ActivePowerScheme: 381b4222-f694-41f0-9685-ff5bb260df2e - So this key seems to be shared for everyone since it is based in HKLM.

So on my machine the "238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0" does not exist since I haven't done anything with that setting, but on the users machine you see: 
reg query 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\Powerschemes\381b4222-f694-41f0-9685-ff5bb260df2e\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0'
DCSettingIndex REG_DWORD 0x0
ACSettingIndex REG_DWORD 0x0

So if it will be needed there can be made a policy to create those 2 keys in that specific one or all the 3 power schema shown which should apply to all the users.

reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\Powerschemes\381b4222-f694-41f0-9685-ff5bb260df2e\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0" /v DCSettingIndex /t REG_DWORD /d 0 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\Powerschemes\381b4222-f694-41f0-9685-ff5bb260df2e\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0" /v ACSettingIndex /t REG_DWORD /d 0 /f
```

## Ghost/Duplicate printers not appearing in control panel

```
Open Notepad:
- Click File, then Print
- Scroll around to find the ghost or duplicate printer(s)
- Select the ghost printer or duplicate printer and RIGHT CLICK
- Select Delete and accept the deletion warning
```

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

## "The computer must be trusted for delegation and the current user account must be configured to allow delegation."

*To work around this problem, set the value of the ProtectionPolicy registry entry to 1 to enable local backup of the MasterKey instead of requiring a RWDC in the following registry subkey.*

```
REG ADD "HKEY_LOCAL_MACHINE\Software\Microsoft\Cryptography\Protect\Providers\df9d8cd0-1501-11d1-8c7a-00c04fc297eb" /v ProtectionPolicy /t REG_DWORD /d 1 /f
```

## "Incorrect permissions on Windows Search directories"

This is a issue where the Windows Search Indexing gets stuck and refuses to rebuild the cache, when running the windows troubleshooter tool it references that the search directories have incorrect permissions. The folder is re-created automatically by the system, so we can just delete it and let it re-generate with the correct rights.

```
Stop-Service "WSearch" -Force
Set-Service -Name "WSearch" -StartupType Disabled -Status Stopped
Remove-Item "C:\ProgramData\Microsoft\Search" -Recurse -Force
Set-Service -Name "WSearch" -StartupType Automatic -Status Running
```

## CSC folder taking up a lot of disk space

In this specific case a user had set the entire shared drive to cache on to the machine, this folder was taking up over 60 GB of disk space.

```
# Disable Offline Files
REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CSC" /v Start /t REG_DWORD /d 4 /f
REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CscService" /v Start /t REG_DWORD /d 4 /f
Set-Service -Name "CscService" -StartupType Disabled -Status Stopped

# Take permissions of cache to clear it
cmd /c 'takeown /f C:\Windows\csc /r /a /d y > NUL'
cmd /c 'icacls C:\Windows\csc /grant Administrators:(F) /t /l /q'

# Clear Cache, using robocopy will purge files with too long path
cmd /c 'mkdir C:\Temp\RoboCSC'
cmd /c 'robocopy "C:\Temp\RoboCSC" "C:\Windows\CSC\v2.0.6\namespace\targetfolder" /purge'
cmd /c 'rmdir "C:\Temp\RoboCSC"'
```

## Shared drives lost after reboot

This seems to be a common issue and especially around the v2004 release of Windows 10. Resetting the registries for the drive mapping appears to be the solution for most of these cases.

```
# First remove all current mapped shared drives!

# Open regedit as the user, browse to the following entries:
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Map Network Drive MRU
- Delete all alphabetic letters containing shared drive paths
HKEY_CURRENT_USER\Network
- Delete all entries in there

# Remap the shared drives again, then go back to:
HKEY_CURRENT_USER\Network
- Click on each entry and make sure that a REG_DWORD with the below exists
* ProviderFlags
* 0x00000001
```

## BSOD - Blue Screen of Death:
### Stop Code: "PAGE_FAULT_IN_NONPAGED_AREA" - "Netwtw06.sys"

*Netwtw06. sys is the Intel WiFi driver*

The issue is solved by updating the Intel WLAN driver!




















---
layout: default
title: Microsoft Office
parent: Documentation
nav_order: 4
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

```
14.0 = Office 2010
15.0 = Office 2013
16.0 = Office 2016

Spreadsheet Compare is only available with Office Professional Plus 2013 or Microsoft 365 Apps for enterprise
Excel: Power Pivot - Requires: Office 2013 professional plus and PowerBI Desktop


get-childitem 'C:\Program Files\Altiris\Altiris Agent\Agents\SoftwareManagement\Software Delivery\' -recurse | where {$_.name -like '*2016*'}
cmd /c 'C:\Program Files\Altiris\Altiris Agent\Agents\SoftwareManagement\Software Delivery\{F734ED03-43DA-4720-AB96-33739378FE9F}\cache\install.bat'
```

```
## To Be Orged:

# Basic .XML file needed for office removal:
# App ID: Access, PrjStd, Visio, Standard, Lync, SkypeforbusinessRetail, ACCESSRT

New-Item -Path 'C:\Packages\office.xml' | add-content -value "<Configuration Product=""Access"">`r`n<Display Level=""none"" CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""Yes"" AcceptEula=""Yes"" />`r`n<Setting Id=""SETUP_REBOOT"" Value=""Never"" />`r`n</Configuration>"

Set-Content 'C:\Packages\office.xml' -value "<Configuration Product=""Access"">`r`n<Display Level=""none"" CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""Yes"" AcceptEula=""Yes"" />`r`n<Setting Id=""SETUP_REBOOT"" Value=""Never"" />`r`n</Configuration>"

cmd /c '"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\setup.exe" /uninstall Standard /config c:\packages\office.xml'
cmd /c '"C:\Program Files\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\setup.exe" /uninstall Standard /config c:\packages\office.xml'

wmic product where "Vendor like '%Microsoft%'" get Name

cmd /c 'TASKKILL /F /IM outlook.exe & TASKKILL /F /IM lync.exe & TASKKILL /F /IM winword.exe & TASKKILL /F /IM excel.exe & TASKKILL /F /IM powerpnt.exe & TASKKILL /F /IM msaccess.exe & TASKKILL /F /IM onenote.exe & TASKKILL /F /IM groove.exe & TASKKILL /F /IM visio.exe & TASKKILL /F /IM winproj.exe & TASKKILL /F /IM MSOSYNC.EXE & TASKKILL /F /IM Teams.exe'

Cscript.exe "C:\Packages\OffScrub_O15msi.vbs” ALL /Quiet /NoCancel /Force /OSE
Cscript.exe "C:\Packages\OffScrub10.vbs” ALL /Quiet /NoCancel /Force /OSE
```

## Office Installation:
### "Error 1907. Could not register font. Verify that you have sufficient permissions to install fonts, and that the system supports this font."
This error is fixed by running "SFC /SCANNOW" which will resolve a file system issue, the command below is intended for a remote powershell session.
```
Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/scannow' -Wait -Verb RunAs -WindowStyle hidden
```

### "Error 1305. Setup cannot read file C:\Program Files (x86)\Common File\Microsoft Shared\OFFICE14\MSO.dll"

The error is related to the cd/dvd drive on the client. The upper filter go between the operating system and the main driver, while the lower driver go between the main driver and the hardware. Don't know exactly what causes the error but removing these keys, then rebooting the machine resolves the issue.

```
reg delete HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E967-E325-11CE-BFC1-08002BE10318} /v LowerFilters /f
reg delete HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E967-E325-11CE-BFC1-08002BE10318} /v UpperFilters /f
```

## Outlook:
### App v16.0.4266.1001 crashing with emails that have attachments
*Microsoft Outlook 2016 may crash when using the Symantec Endpoint Protection (SEP) Outlook Scanner Add-in. Uninstalling or disabling the Symantec add-in resolves the symptoms.*

https://knowledge.broadcom.com/external/article/170414/outlook-2016-crashes-when-using-the-endp.html
- [Office 2013: KB4011282 Download](https://download.microsoft.com/download/6/E/8/6E8FE290-E440-40A4-8F99-5D73CCEF0459/outlook2013-kb4011282-fullfile-x64-glb.exe)
- [Office 2016: KB4011570 Download](https://download.microsoft.com/download/4/B/D/4BDAB50A-A117-45A1-B104-31D946326485/outlook2016-kb4011570-fullfile-x64-glb.exe)

```
Faulting application name: OUTLOOK.EXE, version: 16.0.4266.1001, time stamp: 0x55ba1876
Faulting module name: OUTLOOK.EXE, version: 16.0.4266.1001, time stamp: 0x55ba1876
Exception code: 0xc0000005
Fault offset: 0x000000000002a10f
Faulting process id: 0x26c0
Faulting application start time: 0x01d6ad0dbb0ee84c
Faulting application path: <C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE>
Faulting module path: <C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE>
Report Id: 912141fa-54e5-430c-b8dd-adb8a914bfa7
```

### Tables, images, and attachments are missing in the received meeting requests in Outlook 2013 (sent from Office 2016)
*Assume that a meeting request that has table content, embedded images, and attachments is created in Microsoft Outlook 2016. Then, you receive the meeting request in Outlook 2013. In this situation, the tables, images, and attachments are missing or corrupted in the received meeting request.*

Resolved by installing the patch or editing the user registry:

- [Outlook 2013 KB3127975 Download](https://download.microsoft.com/download/6/1/5/615A5F76-7247-47BC-85BE-EB4CB30A05EF/outlook2013-kb3127975-fullfile-x64-glb.exe)

```
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Options\Calendar" /v AllowHTMLCalendarContent /t REG_DWORD /d 1 /f
```

### Email rule - run a script
This feature was by default hidden after a security update, it can be enabled through the users registry.

```
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\XX.X\Outlook\Security" /v EnableUnsafeClientMailRules /t REG_DWORD /d 1 /f
```

### The Delegates settings were not saved correctly. Unable to activate send-on-behalf-of list. You do not have sufficient permission to perform this operation on this object.

*This error can occur when the Delegates list contains a mailbox user who no longer exists in the organization. To fix the error remove the non-existent user from the Delegates list before you attempt to add other Delegates or change the Delegates settings.*

The first command is to check the attribute, the next following 2 are either to clear a specific inactive user or to remove them all.

```
Get-ADUser $user -Properties publicDelegates | Select-Object -ExpandProperty publicDelegates

Set-ADUser -Remove @{PublicDelegates="UID"}
Set-ADUser -clear PublicDelegates
```

### Items deleted in a shared mailbox goes to the users deleted items instead of the shared mailbox deleted items.

This option can be changed through the user registries. By default the key is set to 8, to store deleted items in the mailbox it was deleted from the key needs to be set to 4.

```
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\XX.0\Outlook\Options\General" /v DelegateWastebasketStyle /t REG_DWORD /d 4 /f
```

### HTTPS linked images in HTML emails display the red X: "The linked image cannot be displayed.  The file may have been moved, renamed, or deleted.  Verify that the link points to the correct file.

This problem occurs when the Internet Explorer Security setting Do not save encrypted pages to disk option is enabled. This can be disabled per user or system wide.

```
reg add "HKEY_CURRENT_USER\software\microsoft\windows\CurrentVersion\Internet Settings" /v DisableCachingOfSSLPages /t REG_DWORD /d 0 /f
reg add "HKEY_LOCAL_MACHINE\software\microsoft\windows\CurrentVersion\Internet Settings" /v DisableCachingOfSSLPages /t REG_DWORD /d 0 /f
```

### Issues when trying to open email attachments

Try to clear the users outlook temp folder and see if that resolves the issue

```
Remove-Item "$env:localappdata\Temporary Internet Files\Content.Outlook\" -confirm:$false -recurse
```

### Search indexing stuck and wont rebuild

By default windows search is responsible for search within outlook, we have had some cases where this have not been functioning properly and with the changes here we can change the search feature over to outlooks own built-in search. This works better for some users in our environment.

```
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v PreventIndexingOutlook /t REG_DWORD /d 1 /f
```

### Email list autocomplete not working

*autocomplete cache has a limit of 1000 entries, after that it starts dropping email addresses off as you add new ones and sometimes fails to add new ones*

The commands will clear the email cache locally and on the server side, then the affected user can test and see how the autofill runs.

```
Remove-Item "$env:userprofile\AppData\Local\Microsoft\Outlook\RoamCache\Stream_Autocomplete*" -confirm:$false
Outlook /cleanautocompletecache
```

*The cache will only repopulate with the addresses you actually send emails to, if you just load users in the "To" field and close the mail then the addresses will stick until you close outlook and then disappear from cache.*

## Excel:
### App v16.0.4978.1000 crashing when saving file to sharepoint

This issue appears with some files that are opened from sharepoint. This issue causes Excel 2016 to crash when you attempt to save the file, so the data gets lost instead of uploaded in to sharepoint.

*Issue appears to be related to the patches KB3191922/KB3203477*

The issue is resolved by installing the appropiate patch:

- [Microsoft Office 2016 - KB3213549 Download](https://download.microsoft.com/download/C/E/3/CE33320D-1400-494F-91E9-14666FD5D979/mso2016-kb3213549-fullfile-x64-glb.exe)

```
Faulting application name: EXCEL.EXE, version: 16.0.4978.1000, time stamp: 0x5e451d6b  
Faulting module name: EXCEL.EXE, version: 16.0.4978.1000, time stamp: 0x5e451d6b  
Exception code: 0xc0000005  
Fault offset: 0x00000000012d9fe9  
Faulting process id: 0x3afc  
Faulting application start time: 0x01d6ce279463d146  
Faulting application path: C:\Program Files\Microsoft Office\Office16\EXCEL.EXE  
Faulting module path: C:\Program Files\Microsoft Office\Office16\EXCEL.EXE  
Report Id: 00135a9e-d418-4adf-9e1c-f27105d52a3b
```

### App v16.0.5026.1000 crashing when saving file with macro
The issue is resolved by installing the appropiate patch or adding a user registry key:

- [Microsoft Excel 2016 - KB3085435 Download](https://download.microsoft.com/download/F/0/5/F05EABEE-70A2-423C-9F91-8AC3B5C65BB3/excel2016-kb3085435-fullfile-x64-glb.exe)
- reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options" /v ForceVBALoadFromSource /t REG_DWORD /d 1 /f

```
Faulting application name: EXCEL.EXE, version: 16.0.5026.1000, time stamp: 0x5ed683b7
Faulting module name: VBE7.DLL, version: 7.1.10.97, time stamp: 0x5ea725f5
Exception code: 0xc0000005
Fault offset: 0x00000000001185ca
Faulting process id: 0x2050
Faulting application start time: 0x01d6c4cb8b20c758
Faulting application path: C:\Program Files\Microsoft Office\Office16\EXCEL.EXE
Faulting module path: C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL
Report Id: faec6dbe-fc1f-4346-9356-12d750d16502
```

### App v16.0.4978.1000 crashing with error reference to chart.dll
The issue is caused by KB4018319 and is solved by installing the appropriate patch:

- [Microsoft Office 2013 - KB2986229 Download](https://download.microsoft.com/download/B/9/9/B99D8FC5-569A-42F0-B0AD-089BB246C327/oart2013-kb2986229-fullfile-x64-glb.exe)
- [Microsoft Office 2016 - KB4011128 Download](https://download.microsoft.com/download/1/A/6/1A69A731-CEC8-4D3F-BA81-0E801E8B0C7D/chart2016-kb4011128-fullfile-x64-glb.exe)

```
Get-EventLog -LogName Application -InstanceId 1000 -Message *EXCEL.exe* | Select-Object -ExpandProperty message
```
```
Faulting application name: EXCEL.EXE, version: 16.0.4978.1000, time stamp: 0x5e451d6b
Faulting module name: chart.dll, version: 16.0.4678.1000, time stamp: 0x5aa7ed63
Exception code: 0xc0000005
Fault offset: 0x00000000001ba0ac
Faulting process id: 0x87c
Faulting application start time: 0x01d6bc44e1f3dcfd
Faulting application path: C:\Program Files\Microsoft Office\Office16\EXCEL.EXE
Faulting module path: C:\Program Files\Microsoft Office\Office16\chart.dll
Report Id: 91239ca4-9421-40c5-9383-ee093ae9cf0e
```

### App v15.0.5233.1000 spreadsheets with macros crashing with error reference to gdi32full.dll

Not sure under what circumstances this issue is triggered but this case was on a Windows 10 v/1809 machine. Issue relates to some windows update released by microsoft, and is resolved by installing a patch based on the OS version.

- [Reference Article](https://kb.parallels.com/en/125027)

- [Windows 10 v1909 - KB4567512 Download](http://download.windowsupdate.com/c/msdownload/update/software/updt/2020/06/windows10.0-kb4567512-x64_2ea636c671529de2154d48a1181c0f02cd919da5.msu)
- [Windows 10 v1809 - KB4567513 Download](http://download.windowsupdate.com/c/msdownload/update/software/updt/2020/06/windows10.0-kb4567513-x64_64dba2ab8d4335fe82a5494bed86f20fd0de3b37.msu)

```
Faulting application name: EXCEL.EXE, version: 15.0.5233.1000, time stamp: 0x5e76d6d3
Faulting module name: gdi32full.dll, version: 10.0.17763.1282, time stamp: 0xf9833b72
Exception code: 0xc0000005
Fault offset: 0x0000000000079de0
Faulting process id: 0x72c
Faulting application start time: 0x01d64df834c8d737
Faulting application path: C:\Program Files\Microsoft Office\Office15\EXCEL.EXE
Faulting module path: C:\WINDOWS\System32\gdi32full.dll
Report Id: 56c03d41-c5b7-4561-952e-2d6a64c7d8df
```

### This file cannot be previewed because there is no previewer installed for it.

This issue is supposedly caused by some Office 2016 patch, there were some machines that had some software changes which caused the fix to never get installed so on these machines the reg key value {00020827-0000-0000-C000-000000000046} did exist, but the registry type was set to REG_EXPAND_SZ and was pointing to a .dll file in the Office16 common files folder. Resolved by creating the key with the proper values

```
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "{00020827-0000-0000-C000-000000000046}" /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers" /v "{00020827-0000-0000-C000-000000000046}" /t REG_SZ /d "Microsoft Excel previewer" /f
```

### Smart View: http session timeouts

The session settings are set per user, this isssue is generally resolved by increasing the timeout time. The changes below is to set the timeout timer to 900000 ms which is 15 minutes

```
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ReceiveTimeout /t REG_DWORD /d 900000 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v KeepAliveTimeout /t REG_DWORD /d 900000 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ServerInfoTimeout /t REG_DWORD /d 900000 /f
```

### Wont open files from shared drive, displays blank page

Have not found an official fix for this issue, but we can work around it by closing all office apps and then renaming the user settings folder. When a office app is then opened then a new settings folder is created with the software defaults.

```
cmd /c 'TASKKILL /F /IM outlook.exe & TASKKILL /F /IM lync.exe & TASKKILL /F /IM winword.exe & TASKKILL /F /IM excel.exe & TASKKILL /F /IM powerpnt.exe & TASKKILL /F /IM msaccess.exe & TASKKILL /F /IM onenote.exe & TASKKILL /F /IM groove.exe & TASKKILL /F /IM visio.exe & TASKKILL /F /IM winproj.exe & TASKKILL /F /IM MSOSYNC.EXE & TASKKILL /F /IM Teams.exe'
Rename-Item "$env:LOCALAPPDATA\Microsoft\Office" "$env:LOCALAPPDATA\Microsoft\Office.old"
```

### Application freeze when copying data from spreadsheet to spreadsheet

Difficult to tell exactly why this happens, it can be when copying a lot or a tiny amount of date from experience. There are a bunch of settings that can be tweaked which i have found to help a lot with the spreadsheet performance.

```
Options > General > User Interface options > enable live preview | untick the option
Options > Advanced > Cut, copy, and paste > Show Paste Options button when content is pasted | untick the option
Options > Advanced > Cut, copy, and paste > Show Insert Options buttons | untick the option
Options > Advanced > Cut, copy, and paste > Cut, copy, and sort inserted objects with their parent cells | untick the option
Options > Advanced > display > disable hardware graphics acceleration | tick the option

Options > Trust Center > Trust Center Settings > Protected View > untick everything
Options > Trust Center > Trust Center Settings > Macro Settings | make sure disable all macros with notification is ticked
Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model | tick the option
Options > Trust Center > Trust Center Settings > ActiveX Settings | make sure prompt me before enabling all controls with minimal restrictions is ticked
```

## OneNote:
### Microsoft OneNote 2013 requires Visual Basic for Applications. If you continue, Visual Basic for Applications will be installed.

This issue is due to some old files being left behind from a previous installation and the installer is unable to override these, so this is why the application tries to configure these on each launch. In this case the client had gone through an upgrade from 2010 to 2013.

```
rename-Item "C:\Program Files\Common Files\Microsoft Shared\VBA" "C:\Program Files\Common Files\Microsoft Shared\VBA.old" -confirm:$false -force

Set-Content 'C:\Packages\office.xml' -value "<Configuration Product=""Standard"">`r`n<Display Level=""none"" CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""Yes"" AcceptEula=""Yes"" />`r`n<Setting Id=""SETUP_REBOOT"" Value=""Never"" />`r`n</Configuration>"

cmd /c '"C:\Program Files\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller\setup.exe" /repair Standard /config c:\packages\office.xml'
```

## Skype for Business
### Webcam freezing after windows update

*After the anniversary update Microsoft dropped support for MJPEG and H264 encoding standards to move towards YUY2 encoding for performance. The change lead to an issue that Webcams that use MJPEG or H264 would not work correctly. This issue then came up again after an later update as well*

To resolve add the registries, then re-open the affected app or restart pc.

```
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Media Foundation\Platform" /v EnableFrameServerMode /t REG_DWORD /d 0 /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows Media Foundation\Platform" /v EnableFrameServerMode /t REG_DWORD /d 0 /f
```

### No emoji icons visible, the IM is just blank with rectangle box

*When users send emojis, they do not see the emoji or receive it from others.

```
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\XX.0\Lync" /v DisableRicherEditCanSetReadOnly /t REG_DWORD /d 1 /f
```

### Precense icons in outlook not working

```
reg add "HKEY_CURRENT_USER\SOFTWARE\IM Providers" /v DefaultIMApp /d Lync /f
```

### Skype meeting button missing in outlook calendar

Resolved by clearing the users office registry key, any manual changes done in the registry to try and force the plugin to load for some reason does not work. This will reset all the user settings for Office.

```
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" /f
```

### App crashing when you click on the toaster notification for any new IM or incoming calls

*Issue appeared after KB4011159 - resolved by installing one of the patches here*

- [Microsoft Office 2016 KB4011631 Download](https://download.microsoft.com/download/E/3/6/E36DCF64-73A4-430E-9E5B-107CE49F714F/msodll40ui2016-kb4011631-fullfile-x64-glb.exe)
- [Microsoft Office 2016 KB4011099 Download](https://download.microsoft.com/download/B/A/1/BA1D3A64-F1AD-49AB-B80E-2D8367D8ACBB/msodll40ui2016-kb4011099-fullfile-x64-glb.exe)

## Office
### Disable Office Sync Tool

The only way to stop this tool from running on boot is to remove the user registries under run.

```
# Disable the sync tool
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /V OfficeSyncProcess /F

# Re-enable the sync tool - 2013
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v OfficeSyncProcess /t REG_SZ /d "\"C:\Program Files\Microsoft Office\Office15\MSOSYNC.EXE"\" /f
```








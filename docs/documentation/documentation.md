---
layout: default
title: Documentation
nav_order: 2
has_children: true
permalink: /docs/documentation
---

# Documentation

This section is a big work in progress to gather much of my quick fixes in a searchable and accessible format.

## RSAT - Remote Server Administration Tools: 

This is a toolkit for IT administrators that allows you to remotely manage roles and features in Windows Server. As of Windows 10 v1809 there is no manual package available to install these tools, so they have to be installed through microsoft. Our environment has a fake wsus endpoint to stop access to windows update so to install the toolkit the commands below needs to be used to set it up.

```
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
Restart-Service wuauserv
Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 1
Restart-Service wuauserv
```


---
layout: default
title: HP Support
parent: Documentation
nav_order: 5
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## HP Machines:
### Intel Management Engine related functionality will be unavailable. Please install the latest version of Intel Management Engine firmware to complete full recovery.

*NOTE: The above message is a false error. There is nothing wrong with the computer, nor the hardware and software installed.*

Issue appears to be related to bios or other hardware settings. One reference article indicates that the issue can be resolved with bios update, this is confirmed further through some posts on the support forums. The second article indicates that potentially there could be issues with the size of the EFI partition.

- [Reference article](https://support.hp.com/my-en/document/c06466020)
- [Reference article](https://support.hp.com/au-en/document/c06466416)

## Docking Station:
### USB-C drivers

- [HP USB-C Universal Dock Driver](https://ftp.hp.com/pub/softpaq/sp92501-93000/sp92798.exe)

### HP USB-C Dock G4, G5 - Resolution Issues on External Monitor(s) While Docked

*The external monitor(s) connected to the dock may exhibit reduced resolution. This issue occurs when the maximum USB-C video throughput has been reached with the BIOS default settings.*

To bypass the following feature needs to be enabled in the BIOS:

- "Enable High Resolution mode when connected to a USB-C DP alt mode dock"

- [Reference article](https://support.hp.com/us-en/document/c06575423)

### HP 2013 UltraSlim Docking Station - Screen resolution issues with 2x DisplayPort's in use

- [Reference Article](https://h30434.www3.hp.com/t5/Business-Notebooks/Firmware-for-UltraSlim-Docking-Stations/td-p/7179965)
- [Firmware Upgrade](https://ftp.hp.com/pub/softpaq/sp79001-79500/sp79015.exe)

### HP EliteBook x360 1030 G2 - Touch Screen Is Not Responding

*The touch screen will not function after re-installing the operating system. This issue occurs with the Wacom AES Digitizer Driver version 7.3.4-30, solution is to install a newer version of the driver.*

```
Wacom AES Digitizer Driver Version: 7.3.4-52
```

### HP EliteDisplay E273q - Monitor Does Not Work Properly When Connected by DisplayPort

- [HP Elite USB-C Dock G3, get firmware version F.16 or later](https://ftp.hp.com/pub/softpaq/sp91501-92000/sp91884.exe)
- [HP USB-C Dock G4, get firmware version F.37 or later](https://ftp.hp.com/pub/softpaq/sp88501-89000/sp88999.exe)















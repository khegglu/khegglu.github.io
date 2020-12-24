---
layout: default
title: Applications
parent: Documentation
nav_order: 3
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## Browsers
### Internet Explorer: This website wants to install the following add-on "crystal report activex viewer control"

> The: ActiveXViewer.cab file needs to be downloaded from the source
> http://yourwebsitewherecrystalisinstalled/crystalreportviewers11/ActiveXControls/ActiveXViewer.cab

Extract ActiveXViewer.cab to a folder on the machine, then manually register the .DLL's, restart browser and access the page again.

```
REGSVR32 /S CRVIEWER.DLL
REGSVR32 /S REPORTPARAMETERDIALOG.DLL
REGSVR32 /S SVIEWHLP.DLL
REGSVR32 /S SWEBRS.DLL
```


### Internet Explorer: DLG_FLAGS_INVALID_CA

This error is due to issues regarding Zscaler's root CA certificate, this needs to be downloaded and installed again as per KB0458729.

- [Zscaler root CA Certificate](http://keyserver.dhl.com/pki/X3/ZscalerRootCertificate-2048-SHA256.crt)
- [IE 11 Instructions](https://help.zscaler.com/zia/configuration-example-importing-zscaler-root-certificate-ie-11)

### Google Chrome: "The application has failed to start because its side-by-side configuration is incorrect."

This error is generally resolved by quickly re-installing google chrome, in our environment there are cases where removing/upgrading chrome fails because the old installation reference is still in the registry. So instead of installing as it should the package delivery system asks for the old installer to remove the old app. In the registry location below there should be a google chrome reference, if it is deleted then the software can be pushed.

```
HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\KeyID
```

## Java
### The following resource is signed with a weak signature algorithm MD5withRSA and is treated as unsigned.

As of java 8u131 applications signed with MD5withRSA/DSA algorithms are treated as unsigned. To bypass this modify the java.security file in the program files folder to still allow the algorithm.

```
((Get-Content -path "C:\Program Files (x86)\Java\jre1.8.*\lib\security\java.security") -replace 'jdk.jar.disabledAlgorithms=MD2, MD5, RSA keySize < 1024','#jdk.jar.disabledAlgorithms=MD2, MD5, RSA keySize < 1024') | Set-Content -Path "C:\Program Files (x86)\Java\jre1.8.*\lib\security\java.security"
```

### Can not verify Deployment Rule Set jar due to certificate expiration

```powershell
remove-item 'C:\Windows\Sun\Java\Deployment\DeploymentRuleSet.jar' -confirm:$false -force
```

### Disable java update prompt



```
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy" /v NotifyDownload /t REG_DWORD /d 0 /f
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy" /v EnableJavaUpdate /t REG_DWORD /d 0 /f
```

## Applications
### Notepad++ "When starting app it is not responding"

This issue was resolved by removing the session.xml from the users profile, then it should work once you open the app again.

```
Remove-Item "C:\Users\$env:username\AppData\Roaming\Notepad++\session.xml" -confirm:$false
```



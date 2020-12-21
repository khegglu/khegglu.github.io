---
layout: default
title: Applications
parent: Documentation
nav_order: 2
---

{: .no_toc }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---

## Browsers
### Internet Explorer: DLG_FLAGS_INVALID_CA

This error is due to issues regarding Zscaler's root CA certificate, this needs to be downloaded and installed again as per KB0458729.

- [Zscaler root CA Certificate](http://keyserver.dhl.com/pki/X3/ZscalerRootCertificate-2048-SHA256.crt)
- [IE 11 Instructions](https://help.zscaler.com/zia/configuration-example-importing-zscaler-root-certificate-ie-11)

### Google Chrome: "The application has failed to start because its side-by-side configuration is incorrect."

This error is generally resolved by quickly re-installing google chrome, in our environment there are cases where removing/upgrading chrome fails because the old installation reference is still in the registry. So instead of installing as it should the package delivery system asks for the old installer to remove the old app. In the registry location below there should be a google chrome reference, if it is deleted then the software can be pushed.

'''
HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\KeyID
'''





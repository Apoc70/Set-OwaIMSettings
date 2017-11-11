# Set-OwaIMSettings.ps1

Validate and update Exchange Server OWA web.config for Skype for Business server name and IM certificate

## Description

This scripts checks Exchange Server 2013+ OWA web.config file for existence of IMCertificateThumbprint and IMServerName Xml nodes required for Skype for Business OWA integration.

IMServerName is the FQDN of the Front End Pool

IMCertificateThumbprint is the certificate thumbprint of the Exchange OWA certificate

This is an enhanced version of Juan Jose Martinez Moreno's script

## Requirements

- Windows Server 2012 R2
- Exchange 2013/2016 Management Shell (aka EMS)
- Identical Exchange setup file path across all Exchange servers
- Identical Exchange server certificate across all Exchange servers having the same thumbprint

## Parameters

### FrontEndPoolFqdn

Skype for Business front end pool FQDN

### CertificateThumbprint

Exchange Server OWA website SSL certificate thumbprint 

## Example

``` PowerShell
.\Set-OwaIMSettings.ps1 -FrontEndPoolFqdn myfepool.varunagroup.de -CertificateThumbprint "1144F22E9E045BF0BA421CAA4BB7AF12EF570C17"
```

Update all OWA web.config files to Skype for Business FE Pool myfepool.varunagroup.de and thumbprint 

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Download and vote at TechNet Gallery

* [https://gallery.technet.microsoft.com/Update-OWA-vDir-config-086bbc17](https://gallery.technet.microsoft.com/Update-OWA-vDir-config-086bbc17)

## Credits

Written by: Thomas Stensitzki

Stay connected:

* My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
* Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
* LinkedIn:	[http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
* Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
* Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
* Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

## Additional Credits:

* Juan Jose Martinez Moreno
* https://juanjosemartinezmoreno.wordpress.com/2015/02/25/modifying-web-config-files-after-exchange-2013-cu-installation/ 

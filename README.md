# Set-OwaIMSettings.ps1
Validate and update Exchange Server OWA web.config for Skype for Business server name and IM certificate

## Description
This scripts checks Exchange Server 2013+ OWA web.config file for existence of IMCertificateThumbprint and IMServerName Xml nodes required for Skype for Business OWA integration.

IMServerName is the FQN of the Front End Pool

IMCertificateThumbprint is the certificate thumbprint of the Exchange OWA certificate

This is an enhanced version of Juan Jose Martinez Moreno's script

## Requirements


## Paramters
### FrontEndPoolFqdn
Skype for Business front end pool FQDN

### CertificateThumbprint
Exchange Server OWA website SSL certificate thumbprint 

## Example
```
.\Set-OwaIMSettings.ps1 -FrontEndPoolFqdn myfepool.varunagroup.de -CertificateThumbprint "1144F22E9E045BF0BA421CAA4BB7AF12EF570C17"
```
Update all OWA web.config files to Skype for Business FE Pool myfepool.varunagroup.de and thumbprint 

## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Update-OWA-vDir-config-086bbc17

## Credits
Written by: Thomas Stensitzki

## Social

* My Blog: http://justcantgetenough.granikos.eu
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog: http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de

## Additional Credits:
* Juan Jose Martinez Moreno
* https://juanjosemartinezmoreno.wordpress.com/2015/02/25/modifying-web-config-files-after-exchange-2013-cu-installation/ 

<# 
  .SYNOPSIS 
  Validate and update Exchange Server OWA web.config for Skype for Business server name and IM certificate

  Thomas Stensitzki 

  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

  Version 1.0, 2017-03-03

  Please send ideas, comments and suggestions to support@granikos.eu 

  .LINK 
  http://scripts.granikos.eu

  .DESCRIPTION 
  This scripts checks Exchange Server 2013+ OWA web.config file for existence 
  of IMCertificateThumbprint and IMServerName Xml nodes required for Skype for Business
  OWA integration.

  IMServerName is the FQN of the Front End Pool

  IMCertificateThumbprint is the certificate thumbprint of the Exchange OWA certificate

  This script is an enhanced version of 
  https://juanjosemartinezmoreno.wordpress.com/2015/02/25/modifying-web-config-files-after-exchange-2013-cu-installation/  
 
  .NOTES 
  Requirements 
  - Windows Server 2012 R2 
  - Exchange 2013/2016 Management Shell (aka EMS)
  - Identical Exchange setup file path across all Exchange servers
  - Identical Exchange server certificate across all Exchange servers having the same thumbprint

    
  Revision History 
  -------------------------------------------------------------------------------- 
  1.0 Initial community release
   
  .PARAMETER FrontEndPoolFqdn
  Skype for Business front end pool FQDN

  .PARAMETER CertificateThumbprint
  Exchange Server OWA website SSL certificate thumbprint 

  .EXAMPLE 
  Update all OWA web.config files to Skype for Business FE Pool myfepool.varunagroup.de and thumbprint 

  .\Set-OwaIMSettings.ps1 -FrontEndPoolFqdn myfepool.varunagroup.de -CertificateThumbprint "1144F22E9E045BF0BA421CAA4BB7AF12EF570C17"
#> 

[CmdletBinding()]
param(
  [string]$FrontEndPoolFqdn = 'myfepool.varunagroup.de',
  [string]$CertificateThumbprint = '1144F22E9E045BF0BA421CAA4BB7AF12EF570C17'
)

# Filepath for UNC- Adjust as needed, if you use any non default path across your servers
# Todo: Fetch local setup path from Exchange server AD object
$ExchangeInstallOwaPath = 'e$\Program Files\Microsoft\Exchange Server\V15\ClientAccess\Owa\'

function Update-WebConfig { 
[CmdletBinding()]
  param(
    [string]$Server = ''
  )

  Write-Host $Server -ForegroundColor Cyan

  # File Backup variables
  $filepath = ('\\{0}\{1}' -f $Server, $ExchangeInstallOwaPath)
  $file = 'web.config'
  $SourceFile = Join-Path -Path $filepath -ChildPath $file
  # Timestamp
  $date = Get-Date -Format 'ddMMyyyyHHmm'
  $BackupFile = Join-Path -Path $filepath -ChildPath ('{0}{1}.bak' -f $file, $date)

  # By default we haven't updated anything
  $Updated = $false

  # Open web.config file in xml format
  $WebConfigFile = ('\\{0}\{1}web.config' -f $Server, $ExchangeInstallOwaPath)
  [xml]$xml = Get-Content -Path $WebConfigFile 

  # Handle IMCertificateThumbprint first
  $IMCertificateThumbprint = $xml.SelectSingleNode("//configuration/appSettings/add[@key='IMCertificateThumbprint']")

  if ($IMCertificateThumbprint -ne $null) {

    # IMCertificateThumbprint exists
    Write-Host (' IMCertificateThumbprint {0} entry found in web.config for server {1}' -f $IMCertificateThumbprint.value, $Server) -ForegroundColor Cyan

    if($IMCertificateThumbprint.value -ne $CertificateThumbprint) {

      # Thumbprtins do not MATCH, so update
      $xml.SelectSingleNode("/configuration/appSettings/add[@key='IMCertificateThumbprint']").value = $CertificateThumbprint

      Write-Host ' IMCertificateThumbprint modified' -ForegroundColor Yellow

      $Updated = $true
    }
    else {
      # Ooops, nothing to do here
      Write-Host ' IMCertificateThumbprint NOT modified' -ForegroundColor Cyan
    }
  }
  else { 
    # IMCertificateThumbprint does NOT exist

    Write-Host (' IMCertificateThumbprint {0} entry NOT found in web.config for server {1}' -f $IMCertificateThumbprint.value, $Server) -ForegroundColor Cyan

    # Create new node
    $root = $xml.get_DocumentElement()
    $IMCertificateThumbprintNode = $xml.CreateNode('element','add','') 
    $IMCertificateThumbprintNode.SetAttribute('key', 'IMCertificateThumbprint')
    $IMCertificateThumbprintNode.SetAttribute('value', $CertificateThumbprint)
    $appSettingsNode = $xml.SelectSingleNode('//configuration/appSettings').AppendChild($IMCertificateThumbprintNode)

    Write-Host ' IMCertificateThumbprint modified' -ForegroundColor Yellow
    $Updated = $true
  }

  # Handle IMServerName first
  $IMServerName = $xml.SelectSingleNode("//configuration/appSettings/add[@key='IMServerName']")

  if ($IMServerName -ne $null) { 

    # IMServerName exists
    Write-Host (' IMServerName {0} entry found in web.config for server {1}' -f $IMServerName.value, $Server) -ForegroundColor Cyan

    if($IMServerName.value -ne $FrontEndPoolFqdn) { 

      # Modifying the entry
      $xml.SelectSingleNode("/configuration/appSettings/add[@key='IMServerName']").value = $FrontEndPoolFqdn

      Write-Host ' IMServerName modified' -ForegroundColor Yellow

      $Updated = $true
    }
    else {
      # Ooops, nothing to do here
      Write-Host ' IMServerName NOT modified' -ForegroundColor Cyan
    }
  }
  else {  
    # IMSerName does NOT exist
    Write-Host (' IMServerName {0} entry NOT found in web.config for server {1}' -f $IMServerName.value, $Server) -ForegroundColor Cyan

    # Create new node
    $root = $xml.get_DocumentElement()
    $IMServerNameNode = $xml.CreateNode('element','add','') 
    $IMServerNameNode.SetAttribute('key', 'IMServerName')
    $IMServerNameNode.SetAttribute('value', $FrontEndPoolFqdn)
    $appSettingsNode = $xml.SelectSingleNode("//configuration/appSettings").AppendChild($IMServerNameNode)
    
    Write-Host ' IMServerName modified' -ForegroundColor Yellow

    $Updated = $true
  }

  if($Updated) { 

    # Create a local copy, if want to sace Xml changes
    Copy-Item -Path $SourceFile -Destination $BackupFile

    # Save file to disk
    $xml.Save($WebConfigFile)
  }
}

### MAIN ########################################

$ExchangeServers = Get-ExchangeServer | Where-Object { $_.AdminDisplayVersion -like '*15.*'} | Sort-Object -Property Name

foreach($ExchangeServer in $ExchangeServers) {
  Update-WebConfig -Server $ExchangeServer.Name
}
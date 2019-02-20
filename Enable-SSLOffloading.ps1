<#
Enable-SSLOffloading
v1.1
12/20/2014
By Nathan O'Bryan
nathan@mcsmlab.com
http://www.mcsmlab.com

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

This script assists in enalbing SSL Offloading for Exchange 2013 SP1 or greater servers.

Change Log
1.1 - Added option to connect to remote Exchange server

#>
Clear-Host

$Answer = Read-Host "Do you want to connect to a remote Exchange server? [Y/N]"

If ($Answer -eq "Y" -or $Answer -eq "y")
    {
    $ExchangeServer = Read-Host "Enter the FQDN of your Exchange server"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session
    }

$MyServer = Read-Host "Type the name of your CA Server (Exchange 2013 SP1 or greater)"
$ExternalHostName = Read-Host "Type the external host name of your CA Server"

Set-OutlookAnywhere -Identity $MyServer\Rpc* -Externalhostname $ExternalHostName -ExternalClientsRequireSsl $True -ExternalClientAuthenticationMethod Basic
Set-OutlookAnywhere -Identity $MyServer\Rpc* -SSLOffloading $true
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS:  -Location "Default Web Site/OWA"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/ecp"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/EWS"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/Autodiscover"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/Microsoft-Server-ActiveSync"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/OAB"
Set-WebConfigurationProperty -Filter //security/access -name sslflags -Value "None" -PSPath IIS: -Location "Default Web Site/MAPI"

iisreset $MyServer
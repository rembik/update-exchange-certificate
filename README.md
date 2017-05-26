# Update MS Exchange certificate
PowerShell script for automation certificate deployment. Tested on Windows Server 2012 R2 with Exchange Version 15.

## Requirements
* PowerShell Version >= 2
* MS Exchange >= 14
* Certificate file [.pfx], private key included

## Usage
```
PS C:\> .\update-exchange-certificate.ps1 [[-PFXPath] <String>] -CertSubject <String> [-PFXPassword <String> ]
                                     [-ExcludeLocalServerCert]
```
All parameters in square brackets are optional. The ExcludeLocalServerCert is forced to $True if left off. You almost never want this set to false, because Exchange server hostname usally isn't equal to the certificate subject why local certificates shouldn't be the updated one.

If the password contains a $ sign, you must escape it with the ` ` character.

### Examples
Install/update certificate in store and activate for Exchange Services:
```
PS C:\> .\update-exchange-certificate.ps1 ".\example.com.pfx" -CertSubject "example.com" -PFXPassword "P@ssw0rd"
```

### Logs
A log file will either be written to %windir%\Temp or to the %LogPath% Task Sequence variable if running from an SCCM\MDT Task.

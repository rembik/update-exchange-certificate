<#

  Syntax examples:
    
    Install/update certificate in store and activate for Exchange Services:
      update-exchange-certificate.ps1 ".\example.com.pfx" -CertSubject "example.com" -PFXPassword "P@ssw0rd"
                                 
    Passing parameters:
      update-exchange-certificate.ps1 [[-PFXPath] <String>] -CertSubject <String> [-PFXPassword <String> ]
                                      [-ExcludeLocalServerCert]
      
    All parameters in square brackets are optional.
    The ExcludeLocalServerCert is forced to $True if left off. You 
    almost never want this set to false, because Exchange
    server hostname usally isn't equal to the certificate 
    subject why local certificates shouldn't be the updated 
    one.

    If the password contains a $ sign, you must escape it with the `
    character.

  Script Name: update-exchange-certificate.ps1
  Release:     1.0
  Written by   brian@rimek.info 19th May 2017
               https://github.com/rembik/update-exchange-certificate

  Note:        This script has been tested thoroughly on Windows 2012R2
               (Exchange 15). I cannot guarantee full
               backward compatibility.

  A log file will either be written to %windir%\Temp or to the
  %LogPath% Task Sequence variable if running from an SCCM\MDT
  Task.

#>

#-------------------------------------------------------------

param (
  [Parameter(Position = 0)][String]$PFXPath,
  [String]$CertSubject=$(throw "Parameter CertSubject is required, please provide a value! e.g. -CertSubject 'example.com'"),
  [String]$PFXPassword,
  [switch]$ExcludeLocalServerCert
)

# Set Powershell Compatibility Mode
Set-StrictMode -Version 2.0

$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($PFXPath)) {
  $PFXPath = $(&$ScriptPath) + "\pkcs12.pfx"
}

if ([String]::IsNullOrEmpty($PFXPassword)) {
  $secPFXPassword = ""
} else {
  $secPFXPassword = ConvertTo-SecureString -String $PFXPassword -Force -AsPlainText
}

if (!($ExcludeLocalServerCert.IsPresent)) { 
  $ExcludeLocalServerCert = $True
}

#-------------------------------------------------------------

Function IsTaskSequence() {
  # This code was taken from a discussion on the CodePlex PowerShell
  # App Deployment Toolkit site. It was posted by mmashwani.
  Try {
      [__ComObject]$SMSTSEnvironment = New-Object -ComObject Microsoft.SMS.TSEnvironment -ErrorAction 'SilentlyContinue' -ErrorVariable SMSTSEnvironmentErr
  }
  Catch {
  }
  If ($SMSTSEnvironmentErr) {
    Write-Verbose "Unable to load ComObject [Microsoft.SMS.TSEnvironment]. Therefore, script is not currently running from an MDT or SCCM Task Sequence."
    Return $false
  }
  ElseIf ($null -ne $SMSTSEnvironment) {
    Write-Verbose "Successfully loaded ComObject [Microsoft.SMS.TSEnvironment]. Therefore, script is currently running from an MDT or SCCM Task Sequence."
    Return $true
  }
}

#-------------------------------------------------------------

$invalidChars = [io.path]::GetInvalidFileNamechars() 
$datestampforfilename = ((Get-Date -format s).ToString() -replace "[$invalidChars]","-")

# Get the script path
$ScriptName = [System.IO.Path]::GetFilenameWithoutExtension($MyInvocation.MyCommand.Path.ToString())
$Logfile = "$ScriptName-$($datestampforfilename).txt"
$logPath = "$($env:windir)\Temp"

If (IsTaskSequence) {
  $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
  $logPath = $tsenv.Value("LogPath")

  $UserDomain = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserDomain")))
  $UserID = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserID")))
  $UserPassword = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserPassword")))
}

$logfile = "$logPath\$Logfile"

# Start the logging 
Start-Transcript $logFile
Write-Output "Logging to $($logFile):"

#-------------------------------------------------------------

Write-Output " +++ Start Certificate Update +++ "

Write-Output " + Loading the Web Administration Module..."
try{
    Import-Module webadministration
}
catch{
    Write-Output " + Failed to load the Web Administration Module!"
}

Write-Output " + Loading the Exchange Management SnapIn..."
try{
    $exchVer =  gcm exsetup | %{$_.fileversioninfo.ProductVersion.Split("{.}")}
    switch ($exchVer[0]) {
        14 {$exchSnapIn = "Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;"}
        15 {$exchSnapIn = "Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;"}		
        default {
            Write-Output " + No supported Exchange version found!"
            exit
        }
    }
    Invoke-Expression $exchSnapIn
}
catch{
    Write-Output " + Failed to load the Exchange Management SnapIn!"
}

if ($exchVer[0] -eq 14 -Or $exchVer[0] -eq 15) {
    Write-Output " + Reading configured Exchange domain names..."
    $exchServer = (Get-ExchangeServer $env:computername).Name
        [array]$exchDomains += ((Get-ClientAccessServer -Identity $exchServer).AutoDiscoverServiceInternalUri.Host).ToLower()  
        [array]$exchDomains += ((Get-OutlookAnywhere -Server $exchServer).ExternalHostname.Hostnamestring).ToLower() 
        [array]$exchDomains += ((Get-OabVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-OabVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
        [array]$exchDomains += ((Get-ActiveSyncVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-ActiveSyncVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
        [array]$exchDomains += ((Get-WebServicesVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-WebServicesVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
        [array]$exchDomains += ((Get-EcpVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-EcpVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
        [array]$exchDomains += ((Get-OwaVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-OwaVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
    if ($exchVer[0] -match 15) {
        [array]$exchDomains += ((Get-OutlookAnywhere -Server $exchServer).Internalhostname.Hostnamestring).ToLower() 
        [array]$exchDomains += ((Get-MapiVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
        [array]$exchDomains += ((Get-MapiVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
    }
    $exchDomains = $exchDomains | select –Unique
    foreach ($exchDomain in $exchDomains) {
        Write-Output "   $exchDomain"
    }
}

Write-Output " + Locating old certificate in store..."
If ($ExcludeLocalServerCert) {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.subject -notlike "CN=$env:COMPUTERNAME"}
} Else {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*"}
}
If ($oldCert) {
    $oldThumbprint = $oldCert.Thumbprint.ToString()
    Write-Output $oldCert
} Else {
    $oldThumbprint = ""
    Write-Output " + Failed to locate old certificate in store!"
}

$imported = $False
Write-Output " + Importing new certificate into Store..."
try{
    $ImportOutput = Import-ExchangeCertificate –FileName $PFXPath -PrivateKeyExportable:$true -Password $secPFXPassword -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable ImportError
    $imported = $True
    write-Output $ImportOutput
}
catch{
    Write-Output " + Failed to import new certificate: $ImportError"
}

If ($imported) {
    Write-Output " + Locating new certificate in store..."
    try{
        If ($ExcludeLocalServerCert) {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint -AND $_.subject -notlike "CN=$env:COMPUTERNAME"}
        } Else {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint}
        }
        $newThumbprint = $newCert.Thumbprint.ToString()
        Write-Output $newCert
    }
    catch{
        Write-Output " + Failed to locate new certificate in store!"
    }

    If ($newCert) {
        Write-Output " + Enable new certificate for Exchange Services SMTP, IMAP, POP and IIS..."
        Enable-ExchangeCertificate -Thumbprint $newThumbprint -Services "SMTP,IMAP,POP,IIS" -Force -ErrorAction SilentlyContinue -ErrorVariable ActivateError
        $checkExchangeThumbprint = (Get-ChildItem -Path IIS:SslBindings | where {$_.port -match "443" -AND $_.IPAddress -match "0.0.0.0" } | select Thumbprint).Thumbprint
        If ($checkExchangeThumbprint -eq $newThumbprint) {
            Write-Output " + Enabled new certificate!"
            If ($oldCert) {
                Write-Output " + Export old certificate as backup and remove it from store..."
                try{
                    If (Test-Path "cert:\LocalMachine\My\$oldThumbprint") {
                        $PFXBackupPath = "$(&$ScriptPath)\$($CertSubject)_backup-$($datestampforfilename).pfx"
                        $ExportOutput = Export-ExchangeCertificate -Thumbprint $oldThumbprint -FileName $PFXBackupPath -Password $secPFXPassword -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable ExportError
                        Enable-ExchangeCertificate -Thumbprint $oldThumbprint -Services "None" -Force -ErrorAction SilentlyContinue
                        Remove-ExchangeCertificate -Thumbprint $oldThumbprint -Confirm:$false
                    }
                }
                catch{
                    Write-Output " + Failed to export and remove old certificate from store: $ExportError"
                }
            }
        } Else {
            Write-Output " + Failed to enable new certificate: $ActivateError"
            If ($oldCert) {
                Write-Output " + Enable old certificate for Exchange Services SMTP, IMAP, POP and IIS..."
                try{
                    Enable-ExchangeCertificate -Thumbprint $oldThumbprint -Services "SMTP,IMAP,POP,IIS" -Force -ErrorAction SilentlyContinue -ErrorVariable ActivateError
                }
                catch{
                    Write-Output " + Failed to enable old certificate: $ActivateError"
                }
            }
            Write-Output " + Remove new certificate from store..."
            try{
                If (Test-Path "cert:\LocalMachine\My\$newThumbprint") {
                    Enable-ExchangeCertificate -Thumbprint $newThumbprint -Services "None" -Force -ErrorAction SilentlyContinue
                    Remove-ExchangeCertificate -Thumbprint $newThumbprint -Confirm:$false
                }
            }
            catch{
                Write-Output " + Failed to remove new certificate from store!"
            }
        }
    }
}

Write-Output " +++ Completed Certificate Update +++ "
 
# Stop logging 
Stop-Transcript

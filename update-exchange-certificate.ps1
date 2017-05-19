<#

  Syntax examples:
    
    Install/update certificate in store and activate for Exchange Services:
      update-exchange-certificate.ps1 ".\example.com.pfx" -CertSubject "example.com" -PFXPassword "P@ssw0rd"
                                 
    Passing parameters:
      update-exchange-certificate.ps1 [[-PFXPath] <String>] -CertSubject <String> [-PFXPassword <String> ]
                                 [-ExcludeLocalServerCert]
      
    All parameters in square brackets are optional.
    The ExcludeLocalServerCert is $True if set. You 
    almost never want this set to true, because Exchange
    server hostname usally is equal to the certificate 
    subject why local certificates could be the updated 
    one. One Exception is using a wildcard certificate. 
    It's there mainly for flexibility.

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

if ($ExcludeLocalServerCert.IsPresent) { 
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
    switch ($exchVer[0])
      {
        14 {$exchSnapIn = "Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;"}
        15 {$exchSnapIn = "Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;"}		
        default
        {
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
  if ($exchVer[0] -match 15)
  {
    [array]$exchDomains += ((Get-OutlookAnywhere -Server $exchServer).Internalhostname.Hostnamestring).ToLower() 
    [array]$exchDomains += ((Get-MapiVirtualDirectory -Server $exchServer).Internalurl.Host).ToLower()
    [array]$exchDomains += ((Get-MapiVirtualDirectory -Server $exchServer).ExternalUrl.Host).ToLower()
  }
  $exchDomains = $exchDomains | select –Unique

  Write-Output " + Configured Exchange domain names:"
  foreach ($exchDomain in $exchDomains) {
    Write-Output "   $exchDomain"
  }
}

Write-Output " + Locating the current(old) certificate in store..."
If ($ExcludeLocalServerCert) {
        $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
    } Else {
        $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*"}
    }
If ($oldCert) {
    $oldThumbprint = $oldCert.Thumbprint.ToString()
    Write-Output " + Current(old) certificate:"
    Write-Output $oldCert
} Else {
    $oldThumbprint = ""
    Write-Output " + Unable to locate current(old) certificate in store!"
}

$ImportSucceed = $False
Write-Output " + Importing certificate into Store..."
try{
    $ImportOutput = Import-PfxCertificate –FilePath $PFXPath -CertStoreLocation "cert:\LocalMachine\My" -Exportable -Password $secPFXPassword -ErrorAction Stop -ErrorVariable ImportError
    $ImportSucceed = $True
    Write-Output " + Imported certificate:"
    write-Output $ImportOutput
}
catch{
    Write-Output " + Failed to import certificate: $ImportError"
}

If ($ImportSucceed) {
    Write-Output " + Locating the new certificate in store..."
    try{
        If ($ExcludeLocalServerCert) {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
        } Else {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint}
        }
        $newThumbprint = $newCert.Thumbprint.ToString()
        Write-Output " + New certificate:"
        Write-Output $newCert
    }
    catch{
        Write-Output " + Unable to locate new certificate in store!"
    }

    If ($newCert) {
        Write-Output " + Activating new certificate for Exchange Services SMTP, IMAP, POP and IIS..."
        try{
            $NewExchangeBinding = Enable-ExchangeCertificate -Thumbprint $newThumbprint -Services "SMTP, IMAP, POP, IIS" –force -ErrorAction SilentlyContinue -ErrorVariable ActivateError
            Write-Output $NewExchangeBinding
            $checkExchangeThumbprint = (Get-ChildItem -Path IIS:SslBindings | where {$_.port -match "443" -AND $_.IPAddress -match "0.0.0.0" } | select Thumbprint).Thumbprint
            If ($checkExchangeThumbprint -eq $newThumbprint) {
                iisreset
                Write-Output " + Activated new certificate!"
            } Else {
                Write-Output " + Unable to activate new certificate: $ActivateError"
            }
        }
        catch{
        }
    }

    If ($oldCert -And $newCert) {
        Write-Output " + Deleting old certificate from Store..."
        try{
            If (Test-Path "cert:\LocalMachine\My\$oldThumbprint") {
                Remove-Item -Path cert:\LocalMachine\My\$oldThumbprint -DeleteKey
            }
        }
        catch{
            Write-Output " + Unable to delete old certificate from store!"
        }
    }
}

Write-Output " +++ Completed Certificate Update +++ "
 
# Stop logging 
Stop-Transcript

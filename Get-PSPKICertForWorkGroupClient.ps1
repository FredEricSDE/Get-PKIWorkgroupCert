##################################################################################
#
# ConfigMgr TS Automated PKI Cert Request for Workgroup Clients
#
# FrameWork Version v.0.1 
# File Version v.0.1 
# Author - Frederic Sellmeier, SVA GmbH
#
# 26.04.2022
#
# Get-PKICertForWorkgroupClient.ps1 - this file. Creates the request based on the 
# MachineName and sends the request to the Issuing CA.
#
# Execution - Create a Package for this file. Add a Task Sequence step for a 
# Powershell script. Add the line under "Example" and set Execution Policy to "Bypass".
#
# Example - .\Get-PSPKICertForWorkgroupClient.ps1 -HostName %_SMSTSMachineName%
#
##################################################################################

# Parameters passed to the script

param(
        [string]$HostName
    )

# Then we need to include all our modules in order to access their functions...

Write-Host "Importing Modules..."

Import-Module .\Modules\GeneralFunctions.psm1 -Global

# Let's start our logging functions to gather as much information as possible

Write-Host "Setting up logging..."

Set-Logfile
Start-Logging

write-LogEntry -Type Information -Message "### Beginning Automated ConfigMgr TS Workgroup PKI Request! ###"

# FQDN of local device and Certificate Template (Template Name NOT Template Display Name!) to be used

$LocalSystemFQDN = ([System.Net.Dns]::GetHostByName(($env:computerName))).HostName
$TemplateName = "<TEMPLATE NAME HERE>" #Template Name NOT Template Display Name!

write-LogEntry -Type Information -Message "Local System FQDN is set to $LocalSystemFQDN"
write-LogEntry -Type Information -Message "Certificate Template is set to $TemplateName"
write-LogEntry -Type Information -Message "

Requesting Workgroup Client Certificate:

Get-Certificate -SubjectName `"CN=$HostName`" ``
-DnsName $LocalSystemFQDN,$HostName -Template $TemplateName ``
-CertStoreLocation Cert:\LocalMachine\My ``
-URL ldap:
"

try{
    Get-Certificate -SubjectName "CN=$HostName" `
    -DnsName $LocalSystemFQDN,$HostName -Template $TemplateName `
    -CertStoreLocation Cert:\LocalMachine\My `
    -URL ldap:

    write-LogEntry -Type Information -Message "Successfully Installed Workgroup Client Certificate!"
}
catch{
    write-LogEntry -Type Error -Message "Requesting Workgroup Client Certificate failed! The following error message was returned:"
    write-LogEntry -Type Error -Message $_.Exception.Message
}

write-LogEntry -Type Information -Message "### Finished Automated ConfigMgr TS Workgroup PKI Request! ###"
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
# GeneralFunctions.psm1 - this file contains general functions required in all
# other modules or files, such as logging.
#
# Execution - this file is not designed to be run seperatly. All functions or a
# superior function are called upon in Get-PSPKICertForWorkgroupClient.ps1.
#
##################################################################################

# This is a test function, to see if the module has been imported correctly into the main programm file

function Test-GeneralFunctions
{
    Write-Host "GenerealFunctions Module imported!"
}

# We use MessageBoxes in the these scripts quite alot. Some Systems don't have MessageBoxes active per default.
# So we need to add this type to the Powershell Assembly

Add-Type -AssemblyName System.Windows.Forms

# Custom .NET class required for CM installation

$source = @"    
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Collections.Generic;

    public static class IniFile
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool WritePrivateProfileString(
		string lpAppName,
		string lpKeyName,
		string lpString,
		string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern uint GetPrivateProfileString(
           string lpAppName,
           string lpKeyName,
           string lpDefault,
           StringBuilder lpReturnedString,
           uint nSize,
           string lpFileName);
		
		[DllImport ("kernel32")]
		private static extern uint GetPrivateProfileString (
			int Section, string Key,
			string Value,
			[MarshalAs (UnmanagedType.LPArray)] byte[] Result, 
			int Size,
			string FileName);
			
		[DllImport ("kernel32")]
		private static extern int GetPrivateProfileString (
			string Section,
			int Key, 
			string Value,
			[MarshalAs (UnmanagedType.LPArray)] byte[] Result,		
			int Size,
			string FileName);
		
        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileSection(
			string lpAppName,
			byte[] lpszReturnBuffer,
			int nSize,
			string lpFileName);

		
        public static void WriteValue(string filePath, string section, string key, string value)
        {
            string fullPath = Path.GetFullPath(filePath);
            bool result = WritePrivateProfileString(section, key, value, fullPath);
        }

        public static string GetValue(string filePath, string section, string key, string defaultValue)
        {
            string fullPath = Path.GetFullPath(filePath);
            var sb = new StringBuilder(1500);
            GetPrivateProfileString(section, key, defaultValue, sb, (uint)sb.Capacity, fullPath);
            return sb.ToString();
        }
		
		public static string[] GetAllCategories(string filePath)
		{
			string fullPath = Path.GetFullPath(filePath);
			for (int maxsize = 25; true; maxsize*=2)
            {
				byte[] buffer = new byte[maxsize];
				uint tempsize = GetPrivateProfileString(0,"","",buffer,maxsize,fullPath);
                int size = (int) tempsize;
				if (size < maxsize -2)
                {
					string tmp = Encoding.ASCII.GetString(buffer,0,size - (size >0 ? 1:0));
					return tmp.Split(new char[] {'\0'});
				}
			}
		}
		public static string[] GetEntryNames(string filePath, string section)
        {
			string fullPath = Path.GetFullPath(filePath);
            for (int maxsize = 25; true; maxsize*=2)
            {
                byte[] bytes = new byte[maxsize];
                int size = GetPrivateProfileString(section,0,"",bytes,maxsize,fullPath);
                if (size < maxsize -2)
                {
                    string entries = Encoding.ASCII.GetString(bytes,0, size - (size >0 ? 1:0));
                    return entries.Split(new char[] {'\0'});
                }
            }
        }
		public static List<string> GetAllValuesOfCategory(string iniFileMulti, string category)
		{
			byte[] buffer = new byte[2048];

			GetPrivateProfileSection(category, buffer, 2048, iniFileMulti);
			string[] tmp = Encoding.ASCII.GetString(buffer).Trim('\0').Split('\0');

			List<string> result = new List<string>();

			foreach (string entry in tmp)
			{
				result.Add(entry.Substring(0, entry.IndexOf("=")));
			}

			return result;
		}		
    }
"@

# Adds the custom Microsoft .NET Framework type (a class) to the Windows PowerShell session

Add-Type -TypeDefinition $source

# The logging functionality has been imported from the TechNet Gallery (https://gallery.technet.microsoft.com/scriptcenter/Powershell-Logging-Module-fbacdffd)

<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.128
	 Created on:   	7/09/2016 10:55 AM
	 Created by:   	M Houston
	 Organization: 	Camden Council
	 Filename: PSLogging.psm1
	===========================================================================
	.DESCRIPTION
		A logging module for Powershell that:

		1. Creates a log file in the repository based on the script name
		2. Archives old log files
		3. Logs in CMTRACE format for easy visible parsing
#>

#region Variables
$scriptName = GCI $MyInvocation.PSCommandPath | Select -Expand Name
$logFileName = "$scriptName.log"
$logPathName = "C:\Windows\Temp" #Split-Path $script:MyInvocation.MyCommand.Path
$logPathName = "$logPathName\Logs\"
$logFullPath = "$($logPathName)$($logFileName)"
$logSize = "5MB"
$logCount = 5
#endregion

function Set-LogFile
{
	param(
        [string]$LogPath
    )

    <#
		Checks if the log file exists, archives if required.
	#>

	# Check log exists
	If (!(Test-Path "$logFullPath"))
	{
		# Check path exists
		If (!(Test-Path "$logPathName"))
		{
			New-Item -Path "$logPathName" -Type Directory
		}
		New-Item -Path "$logFullPath" -Type File
	}
	else
	{
		# Check log size
		if ((Get-Item "$logFullPath").length -gt $logSize)
		{
			#Archive the completed log
			Move-Item -Path "$logFullPath" -Destination "$logPathName\$(Get-Date -format yyyyMMdd)-$(Get-Date -Format HHmmss)-$($logFileName -replace "ps1", "bak")"
			#Create new Log
			New-Item -Path "$logFullPath" -Type File
			
		}
	}
	#Check number of Archives
	While ((Get-ChildItem "$logPathName\*.bak.log").count -gt $logCount)
	{
		Get-ChildItem "$logPathName\*.bak.log" | Sort CreationTime | Select -First 1 | Remove-Item
	}
}

function Write-LogEntry
{
	<#
		Writes a single line to the log file in CMTRACE format.
		Uses either 'Error','Warning' or 'Information' to set
		the visibility in the log file.
	#>
	#Define and validate parameters 
	[CmdletBinding()]
	Param (
		#The information to log 
		[parameter(Mandatory = $True)]
		$Message,
		#The severity (Error, Warning, Verbose, Debug, Information)

		[parameter(Mandatory = $false)]
		[ValidateSet('Warning', 'Error', 'Verbose', 'Debug', 'Information')]
		[String]$Type = "Information",
		#Write back to the console or just to the log file. By default it will write back to the host.

		[parameter(Mandatory = $False)]
		[switch]$WriteBackToHost = $True
		
	) #Param
	
	#Get the info about the calling script, function etc
	$callinginfo = (Get-PSCallStack)[1]
	
	#Set Source Information
	$Source = (Get-PSCallStack)[1].Location
	
	#Set Component Information
	$Component = (Get-Process -Id $PID).ProcessName
	
	#Set PID Information
	$ProcessID = $PID
	
	#Obtain UTC offset 
	$DateTime = New-Object -ComObject WbemScripting.SWbemDateTime
	$DateTime.SetVarDate($(Get-Date))
	$UtcValue = $DateTime.Value
	$UtcOffset = $UtcValue.Substring(21, $UtcValue.Length - 21)
	
	#Set the order 
	switch ($Type)
	{
		'Warning' { $Severity = 2 } #Warning
		'Error' { $Severity = 3 } #Error
		'Information' { $Severity = 6 } #Information
	} #Switch
	
	switch ($severity)
	{
		2{
			#Warning
			
			#Write the log entry in the CMTrace Format.
			$logline = `
			"<![LOG[$($Type.ToUpper()) - $message.]LOG]!>" +`
			"<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
			"date=`"$(Get-Date -Format M-d-yyyy)`" " +`
			"component=`"$Component`" " +`
			"context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
			"type=`"$Severity`" " +`
			"thread=`"$ProcessID`" " +`
			"file=`"$Source`">";
			$logline | Out-File -Append -Encoding utf8 -FilePath ('FileSystem::' + $logFullPath);
			
		} #Warning
		
		3{
			#Error
			
			#Write the log entry in the CMTrace Format.
			$logline = `
			"<![LOG[$($Type.ToUpper()) - $message.]LOG]!>" +`
			"<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
			"date=`"$(Get-Date -Format M-d-yyyy)`" " +`
			"component=`"$Component`" " +`
			"context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
			"type=`"$Severity`" " +`
			"thread=`"$ProcessID`" " +`
			"file=`"$Source`">";
			$logline | Out-File -Append -Encoding utf8 -FilePath ('FileSystem::' + $logFullPath);
			
		} #Error
		
		6{
			#Information
			
			#Write the log entry in the CMTrace Format.
			$logline = `
			"<![LOG[$($Type.ToUpper()) - $message.]LOG]!>" +`
			"<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
			"date=`"$(Get-Date -Format M-d-yyyy)`" " +`
			"component=`"$Component`" " +`
			"context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
			"type=`"$Severity`" " +`
			"thread=`"$ProcessID`" " +`
			"file=`"$Source`">";
			$logline | Out-File -Append -Encoding utf8 -FilePath ('FileSystem::' + $logFullPath);
			
		} #Information
	}
}

function Start-Logging
{
	<#
		Parent function for future expansion of functionality
	#>
	Set-LogFile
	Write-LogEntry -Message "[Starting Script Execution]`r`r`n User Context : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`r`r`n Hostname : $($env:COMPUTERNAME)"
}
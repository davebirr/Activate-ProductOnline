<#
========================================================================
	File:		Activate-ProductOnline.ps1
	Version:	0.16.1
	Author:		David Bjurman-Birr, adapting work from Daniel Dorner's ActivateWs project on GitHub
	Date:		01/14/2020
	
	Purpose:	Installs and activates a product key
	
	Usage:		./Activate-ProductOnline.ps1 <Required Parameter> [Optional Parameter]
	
                <-ProductKey>          <Specifies the product key>
                [-LogFile]             [Specifies the full path to the log file]
	
	This script code is provided "as is", with no guarantee or warranty concerning
	the usability or impact on systems and may be used, distributed, and
	modified in any way provided the parties agree and acknowledge the 
	Microsoft or Microsoft Partners have neither accountability or 
	responsibility for results produced by use of this script.
	
	Microsoft will not provide any support through any means.

========================================================================
#>

param (
	[Parameter(
		Mandatory = $true,
		ValueFromPipeline = $true,
		HelpMessage = 'Specifies the product key. It is a 25-character code and looks like this: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
		Position = 0)]
	[ValidatePattern('^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$')]
	[string]$ProductKey,

	[Parameter(
		Mandatory = $false,
		ValueFromPipeline = $true,
		HelpMessage = 'Specifies the full path to the log file, e.g. "C:\Log\Logfile.log"',
		Position = 2)]
	[ValidateNotNullorEmpty()]
	[string]$LogFile = "$env:TEMP\Activate-ProductOnline.log"
)

function LogAndConsole($Message)
{
	try {
		if (!$logInitialized) {
			"{0}; <---- Starting {1} on host {2}  ---->" -f (Get-Date), $MyInvocation.ScriptName, $env:COMPUTERNAME | Out-File -FilePath $LogFile -Append -Force
			"{0}; {1} version: {2}" -f (Get-Date), $script:MyInvocation.MyCommand.Name, $scriptVersion | Out-File -FilePath $LogFile -Append -Force
			"{0}; Initialized logging at {1}" -f (Get-Date), $LogFile | Out-File -FilePath $LogFile -Append -Force
			
			$script:logInitialized = $true
		}
		
		foreach ($line in $Message) {
			$line = "{0}; {1}" -f (Get-Date), $line
			$line | Out-File -FilePath $LogFile -Append -Force
		}
		
		Write-Host $Message
		
	} catch [System.IO.DirectoryNotFoundException] {
		$script:LogFile = "$env:TEMP\Activate-ProductOnline.log"
		Write-Host "[Warning] Could not find a part of the path $LogFile. The output will be redirected to $LogFile." 
		
	} catch [System.UnauthorizedAccessException] {
		$script:LogFile = "$env:TEMP\Activate-ProductOnline.log"
		Write-Host "[Warning] Access to the path $LogFile is denied. The output will be redirected to $LogFile."
		
	} catch {
		Write-Host  "[Error] Exception calling 'LogAndConsole':" $_.Exception.Message
		Exit $MyInvocation.ScriptLineNumber
	}
}

function InstallAndActivateProductKey([string]$ProductKey) {
    #Install Key, equivalent to slmgr /ipk
	try {
		# Check if product key is already installed and activated.
		$partialProductKey = $ProductKey.Substring($ProductKey.Length - 5)
		$licensingProduct = Get-WmiObject -Query ('SELECT LicenseStatus FROM SoftwareLicensingProduct where PartialProductKey = "{0}"' -f $partialProductKey)
		
		if ($licensingProduct.LicenseStatus -eq 1) {
			LogAndConsole "The product is already activated."
			Exit $MyInvocation.ScriptLineNumber
		}
	
		# Install the product key.
		LogAndConsole "Installing product key $ProductKey ..."
		$licensingService = Get-WmiObject -Query 'SELECT VERSION FROM SoftwareLicensingService'
        LogAndConsole "The Software Licensing Server Version is $licensingService.Version"
		$licensingService.InstallProductKey($ProductKey) | Out-Null
		$licensingService.RefreshLicenseStatus() | Out-Null

	} catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
		LogAndConsole "[Error] Failed to install the product key."
        LogAndConsole "[Error] The error message was: $ErrorMessage"
		Exit $MyInvocation.ScriptLineNumber
	}

    #Activate Key
	try {
		# Get the licensing information.
		LogAndConsole "Retrieving license information..."
		$licensingProduct = Get-WmiObject -Query ('SELECT ID, Name, OfflineInstallationId, ProductKeyID FROM SoftwareLicensingProduct where PartialProductKey = "{0}"' -f $partialProductKey)

		if(!$licensingProduct) {
			LogAndConsole "No license information for product key $ProductKey was found."
			Exit $MyInvocation.ScriptLineNumber
		}
		
		$licenseName = $licensingProduct.Name                       # Name  
		$InstallationId = $licensingProduct.OfflineInstallationId   # Installation ID
		$activationId = $licensingProduct.ID                        # Activation ID
		$ExtendedProductId = $licensingProduct.ProductKeyID         # Extended Product ID
	   
		LogAndConsole "Name             : $licenseName"
		LogAndConsole "Installation ID  : $InstallationId"
		LogAndConsole "Activation ID    : $activationId"
		LogAndConsole "Extd. Product ID : $ExtendedProductId"
		
	} catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
		LogAndConsole "[Error] Failed to retrieve the license information. $FailedItem"
        LogAndConsole "[Error] The error message was $ErrorMessage"
		Exit $MyInvocation.ScriptLineNumber
	}

	try {
		# Activate the product using online service
		LogAndConsole "Activating product..."
		#$licensingProduct.DepositOfflineConfirmationId($InstallationId, $confirmationId) | Out-Null
        $licensingProduct.Activate() | Out-Null
		$licensingService.RefreshLicenseStatus() | Out-Null
		
		# Check if the activation was successful.
		$licensingProduct = Get-WmiObject -Query ('SELECT LicenseStatus, LicenseStatusReason FROM SoftwareLicensingProduct where PartialProductKey = "{0}"' -f $partialProductKey)
		
		if (!$licensingProduct.LicenseStatus -eq 1) {
			LogAndConsole "[Error] Product activation failed ($($licensingProduct.LicenseStatusReason))."
			Exit $MyInvocation.ScriptLineNumber
		}
		
		LogAndConsole "Product activated successfully."
		
	} catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
		LogAndConsole "[Error] Failed to activate product."
        LogAndConsole "[Error] The error message was $ErrorMessage"
		Exit $MyInvocation.ScriptLineNumber
	}
}

function Main {
	LogAndConsole ""
	InstallAndActivateProductKey($ProductKey)
}

Main
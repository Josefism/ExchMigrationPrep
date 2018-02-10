# Script to get Exchange Attributes and Proxy Addresses out of AD prior to migration
[CmdletBinding()]
Param(
	[Parameter(Position=1)]
	# Indicates the type of operation being performed.
	# GET operation extracts data prior to migration.
	# PUT operation inserts data after migration.
	[string]$operationType,
	
	[Parameter(Position=2)]
	# Accepts the path to the working directory where export files 
	# are created. Do NOT enclose the path in quotes. Begin with drive letter
	# and do not include a filename. Trailing slash is optional.
	# EX: C:\path\to\working\directory
	[string]$pathToWorkDir,

	[Parameter(Position=3)]
	# Accepts the Search Base for the Active Directory query. 
	# Do NOT enclose in quotes. DO include both OU and DC elements.
	# EX: OU=Monroe,OU=NetTech LLC,DC=NT,DC=lan
	[string]$searchBase	
)

# Capture current date-time for logging
$runDateTime = Get-Date -format "yyyy-MM-dd @ hh:mm"

# Keep track of export file naming in case of multiple script executions
$fileDateStamp = Get-Date -format "yyyy-MM-dd"

# Set variable for script operations logging filename
$operationLogFile = "ExchMigrationPrepLog_" + $fileDateStamp + "_1.log"

# Begin logging script operations
$operationLog = "# === Exchange Migration Data Script Log === #`r`n"
$operationLog += "Operations Started: " + $runDateTime + "`r`n`r`n"

# Begin UI Feedback
Write-Host "`n# === Exchange Migration Data Script Running === #`n"

# Check for Active Directory module in PowerShell to see how to proceed with GET operations
$psADModPresent = $True
if (!(Get-Module -ListAvailable -Name "ActiveDirectory")) {
	$psADModPresent = $False
	$operationLog += "Active Directory PowerShell Modules: Unavailable`r`n`r`n"
	Write-Host "Active Directory Modules: Unavailable - GET Output will be in compact file"
} else {
	$operationLog += "Active Directory PowerShell Modules: Available`r`n`r`n"
	Write-Host "Active Directory Modules: Available - GET Output will be in separate files"
	Import-Module ActiveDirectory
}

# If parameters were not provided, prompt for each one (makes neater prompts than "Mandatory" parameter flag)

# Variables for quitting entire script at user prompts - check to cancel operations at every user prompt
$quitNow = $False
$quitAll = [System.Management.Automation.Host.ChoiceDescription]::new("&Quit")
$quitAll.HelpMessage = "Cancel Exchange Migration Prep and exit without continuing"	

####################
# PATH TO WORK DIR #
####################
# Establish a working directory for script input/output 
# This directory is the home for receiving data exports, providing data for imports, and logging script actions
if ($pathToWorkDir -eq '') {
	Write-Host "`nThis script requires a working directory for file output."
	Write-Host "Enter a path, including drive letter. Do not include filename."
	Write-Host "If path does not exist on specified drive it will be created.`n"
	
	do {
		try {
			$driveExists = $False
			$pathInput = Read-Host -Prompt "Path to working directory"
			
			# If user entered path wrapped in quotes, remove the quotes
			if ($pathInput -match '^["'']') { $pathInput = $pathInput -replace '["'']', '' }

			# Correct forward slashes to backslashes if needed
			if ($pathInput -match '[/]') { $pathInput = $pathInput -replace '[/]', '\' }
			
			# Check the entered path for trailing backslash and add if not present
			if ($pathInput -notmatch '.+?\\$') { $pathInput += "\"}

			# Ensure user entered a valid drive letter that is mapped on the system
			# Path should begin with single letter followed by colon
			# Letter provided must resolve to a mapped drive
			if ($pathInput -match '^[A-z]:') { 
				$driveInput = $pathInput.Substring(0,1) 
				if (Get-PSDrive | Where { $_.Name -eq $driveInput }) {
					$driveExists = $True
				} else {
					$driveExists = $False
					Write-Host "`nEntered drive letter not found.`n" -ForegroundColor Red
				}
			} else {
				$driveExists = $False
				Write-Host "`nPath format incorrect. Begin with drive letter followed by colon and backslash. EX: C:\path\to\directory`n" -ForegroundColor Red
			}
		}
		catch {
			$driveExists = $False
		}
	} until ( $driveExists -eq $True )
	
	# Once we have a valid path input, assign it to our variable
	$pathToWorkDir = $pathInput
	
	$operationLog += "Path to Work Directory: " + $pathToWorkDir + " [From Prompt]`r`n`r`n"
} else {
	# Check the pathToWorkDir parameter for proper formatting, prompt for corrections if needed

	# If user entered path wrapped in quotes, remove the quotes
	if ($pathToWorkDir -match '^["'']') { $pathToWorkDir = $pathToWorkDir -replace '["'']', '' }
	
	# Correct forward slashes to backslashes if needed
	if ($pathToWorkDir -match '[/]') { $pathToWorkDir = $pathToWorkDir -replace '[/]', '\' }

	# Check the pathToWorkDir parameter for trailing backslash and add if not present
	if ($pathToWorkDir -match '.+?\\$') { continue } else { $pathToWorkDir += "\"}

	# Variables for validation testing
	$driveParamExists = $True
	
	# Ensure parameter contains a valid drive letter that is mapped on the system
	# Path should begin with single letter followed by colon
	# Letter provided must resolve to a mapped drive
	if ($pathToWorkDir -match '^[A-z]:') { 
		$driveParam = $pathToWorkDir.Substring(0,1) 
		if (Get-PSDrive | Where { $_.Name -eq $driveParam }) {
			$driveParamExists = $True
			$operationLog += "Path to Work Directory: " + $pathToWorkDir + " [From Parameters]`r`n`r`n"
			Write-Host "`nPath to Work Directory: $pathToWorkDir - Determined from parameters`n"
		} else {
			$driveParamExists = $False
		}
	} else {
		$driveParamExists = $False
	}
	
	# If path from parameter contains invalid drive letter or bad path format, prompt for path correction
	if ($driveParamExists -eq $False) {
		Write-Host "`nThe provided Path to Working Directory parameter was invalid." -ForegroundColor Red
		Write-Host "Please provide a valid path to working directory now."
		Write-Host "Enter a path, including drive letter. Do not include filename."
		Write-Host "If path does not exist on specified drive it will be created.`n"

		do {
			try {
				$driveCorrectedExists = $True
				$correctedPathInput = Read-Host -Prompt "Path to working directory"
							
				# If user entered path wrapped in quotes, remove the quotes
				if ($correctedPathInput -match '^["'']') { $correctedPathInput = $correctedPathInput -replace '["'']', '' }
				
				# Correct forward slashes to backslashes if needed
				if ($correctedPathInput -match '[/]') { $correctedPathInput = $correctedPathInput -replace '[/]', '\' }

				# Check the pathToWorkDir parameter for trailing backslash and add if not present
				if ($correctedPathInput -match '.+?\\$') { continue } else { $correctedPathInput += "\"}

				# Ensure entered correction contains a valid drive number that is mapped on the system
				# Path should begin with single letter followed by colon
				# Letter provided must resolve to a mapped drive
				if ($correctedPathInput -match '^[A-z]:') { 
					$driveCorrected = $correctedPathInput.Substring(0,1) 
					if (Get-PSDrive | Where { $_.Name -eq $driveCorrected }) {
						$driveCorrectedExists = $True
						$pathToWorkDir = $correctedPathInput
						$operationLog += "Path to Work Directory: " + $pathToWorkDir + " [From Correction Prompt]`r`n`r`n"
					} else {
						$driveCorrectedExists = $False
						Write-Host "`nEntered drive letter not found." -ForegroundColor Red
					}
				} else {
					$driveCorrectedExists = $False
					Write-Host "`nPath format incorrect. Begin with drive letter followed by colon and backslash. EX: C:\path\to\directory" -ForegroundColor Red
				}
			}
			catch {
				$driveCorrectedExists = $False
			}
		} until ( $driveCorrectedExists -eq $True )	
	}
}

# Once drive letter and path format are validated, we need to create the path if it doesn't already exist
if ($pathToWorkDir | Test-Path) {
	# Set variables to handle zero sum instances
	$attributesExportNum = "0"
	$proxiesExportNum = "0"
	$compactExportNum = "0"
	
	# If the path exists, check for existing files so we can adjust export file names if needed
	if (Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedAttributes*") {
		$attributesExportNum = $(Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedAttributes*").Count
		$attributesFileName = "ExportedAttributes_" + $fileDateStamp + "_" + $($attributesExportNum + 1) + ".csv"
	}
	
	if (Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedProxies*") {
		$proxiesExportNum = $(Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedProxies*").Count
		$proxiesFileName = "ExportedProxies_" + $fileDateStamp + "_" + $($attributesExportNum + 1) + ".csv"
	}
	
	if (Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedCompact*") {
		$compactExportNum = $(Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExportedCompact*").Count
		$compactFileName = "ExportedCompact_" + $fileDateStamp + "_" + $($compactExportNum + 1) + ".csv"
	}	

	if (Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExchMigrationPrepLog*") {
		$operationLogNum = $(Get-ChildItem -path $pathToWorkDir -Recurse -Filter "ExchMigrationPrepLog*").Count
		$operationLogFile = "ExchMigrationPrepLog_" + $fileDateStamp + "_" + $($operationLogNum + 1) + ".log"
	}

	$operationLog += "Working Directory Found`r`n"
	$operationLog += "Existing Attributes Export files: " + $attributesExportNum + "`r`n"
	$operationLog += "Existing Proxies Export files: " + $proxiesExportNum + "`r`n"
	$operationLog += "Existing Compact Export files: " + $compactExportNum + "`r`n`r`n"
	
	Write-Host "`nWorking Directory Found"
	Write-Host "Existing Attributes Export files: $attributesExportNum"
	Write-Host "Existing Proxies Export files: $proxiesExportNum"
	Write-Host "Existing Compact Export files: $compactExportNum"
} else {
	# If the path does not exist, create it
	$dirCreateSuccess = $False
	try {
		New-Item -ItemType directory -Path $pathToWorkDir -ea stop
		$dirCreateSuccess = $True
	}
	catch {
		# TODO: Consider adding logic to loop user back through working directory selection
		Write-Host "`nProblem creating directory at the path provided.`nPlease run Exchange Migration Prep again using a different path for working directory.`n`n" -ForegroundColor Red
		exit
	}

	if ($dirCreateSuccess) {
		# Log successful creation
		$operationLog += "Working Directory Created`r`n`r`n"
		Write-Host "`nWorking Directory Created"
	}
}

# Define the path to the operation log for this script
$operationLogOutput = $pathToWorkDir + $operationLogFile
	
##################
# OPERATION TYPE #
##################
# Determine OperationType (GET/PUT) to decide what portions of the script to execute
# If operationType parameter was not specified, prompt user for operation type selection
if ($operationType -eq '') {
	$operationTitle = "Select Operation Type"
	$operationInfo = "Are you getting data prior to migration or putting data after migration? Default is GET"
	$opGet = [System.Management.Automation.Host.ChoiceDescription]::new("&Get")
	$opGet.HelpMessage = "Get Exchange data out of Active Directory prior to migration"
	$opPut = [System.Management.Automation.Host.ChoiceDescription]::new("&Put")
	$opPut.HelpMessage = "Put Exchange data into Active Directory after to migration"
	$operationSelectOptions = [System.Management.Automation.Host.ChoiceDescription[]] @($opGet, $opPut, $quitAll)
	[int]$operationDefault = 0
	$operationChoice = $host.UI.PromptForChoice($operationTitle, $operationInfo, $operationSelectOptions, $operationDefault)
	switch($operationChoice) {
		0 { $operationType = "GET"; break }
		1 { $operationType = "PUT"; break }
		2 { $quitNow = $True; break }
	}
	
	if ($quitNow) {
		# Capture current date-time for logging
		$endDateTime = Get-Date -format "yyyy-MM-dd @ hh:mm"
		
		$operationLog += "Exchange Migration Prep cancelled by user at Operation Type prompt.`r`n"
		$operationLog += "No output files generated.`r`n`r`n"
		$operationLog += "Exchange Migration Prep ended at " + $endDateTime + "`r`n"

		$operationLog | Out-File $operationLogOutput
				
		Write-Host "`nExchange Migration Prep cancelled by user. Log file saved to working directory:" -ForegroundColor DarkCyan
		Write-Host $operationLogOutput -ForegroundColor Green
		Write-Host "`n"
		
		exit
	} else {
		$operationLog += "Operation Type: " + $operationType + " [From Prompt]`r`n`r`n"
	}	
} else {
	# Standardize OperationType parameter, or apply default if no match 
	$operationType = $operationType.ToUpper()
	
	# Handle partial parameter entries
	switch($operationType) {
		"G" { $operationType = "GET"; break }
		"GE" { $operationType = "GET"; break }
		"GET" { $operationType = "GET"; break }
		"P" { $operationType = "PUT"; break }
		"PU" { $operationType = "PUT"; break }
		"PUT" { $operationType = "PUT"; break }
		default { $operationType = "GET"; break }
	}
	
	$operationLog += "Operation Type: " + $operationType + " [From Parameters]`r`n`r`n"
	Write-Host "Operation Type: $operationType - Determined from parameters`n`n"
}

# Set filename variables for output
$attributesOutput = $pathToWorkDir + $attributesFileName
$proxiesOutput = $pathToWorkDir + $proxiesFileName
$compactOutput = $pathToWorkDir + $compactFileName

# Discover Distinguished Names of all possible OUs
# Used for either parameter validation or menu-based selection if parameter not provided
$i = 0
$ouDiscoveries = @{}
$ouChoices = @()
$ouSelectOptions = [System.Management.Automation.Host.ChoiceDescription[]] @()
$ouSelectHelpers = "`nDiscovered Search Base Options:"
$ouInfo = ([adsisearcher]"objectcategory=organizationalunit")
$ouInfo.PropertiesToLoad.Add("DistinguishedName")
$ouInfo.findAll() | ForEach-Object { $ouDiscoveries.Add( "&$i", $_.Properties["DistinguishedName"] ); $i += 1 }
$ouDiscoveries.GetEnumerator() | Sort-Object -Property Name | 
  ForEach-Object { 
	$ouChoices += , @($_.Name, $_.Value)
	$ouOpt = [System.Management.Automation.Host.ChoiceDescription]::new($($_.Name))
	$ouOpt.HelpMessage = $_.Value
	$ouSelectOptions += $ouOpt
	$ouSelectHelpers += $("`n[" + $_.Name.Substring(1) + "] " + $_.Value) 
  }

# Include option to cancel entire script
$ouSelectHelpers += "`n[Q] Quit"
$ouSelectOptions += $quitAll
  
$ouSelectTitle = "Select Search Base"
$ouSelectInfo = "Choose OU from options listed above (Default is 0)"
$ouSelectDefault = 0

###############
# SEARCH BASE #
###############
# Verify search base parameter if provided, else, prompt for selection

# Search base validity flag
$searchBaseReady = $False

# If searchBase parameter was provided, check through found OUs to see if it is valid
if ($searchBase -ne '') {
	foreach ($ou in $ouChoices) { 
		if ($ou -match $( "^" + $searchBase + "$")) {
			$searchBaseReady = $True
	
			$operationLog += "Search Base Parameter Validated: " + $searchBase + "`r`n"	
			Write-Host "`nSearch Base Parameter Validated: $searchBase"
			break
		} else {
			$ouSelectInfo = "Provided searchBase parameter was not valid. Choose OU from options listed above (Default is 0)"
			$operationLog += "Provided Search Base Parameter was invalid - Prompting for selection`r`n"
		}
	}
}

# If searchBase parameter was NOT provided, or if provided searchBase was invalid, prompt for selection
if ($searchBaseReady -eq $False) {
	Write-Host $ouSelectHelpers
	$searchBaseSelection = $Host.UI.PromptForChoice( $ouSelectTitle, $ouSelectInfo, $ouSelectOptions, $ouSelectDefault )
	
	# Process search base selection unless user chose option to quit
	if ($searchBaseSelection -ne "Q") {
		$searchBase = $ouChoices[$searchBaseSelection][1]
		
		$operationLog += "Search Base Selection Validated: " + $searchBase + "`r`n"
		Write-Host "`nSearch Base Selection Validated: $searchBase"
	} else {
		# Capture current date-time for logging
		$endDateTime = Get-Date -format "yyyy-MM-dd @ hh:mm"
		
		$operationLog += "Exchange Migration Prep cancelled by user at Search Base prompt.`r`n"
		$operationLog += "No output files generated.`r`n"
		$operationLog += "Exchange Migration Prep ended at " + $endDateTime + "`r`n"

		$operationLog | Out-File $operationLogOutput
				
		Write-Host "`nExchange Migration Prep cancelled by user. Log file saved to working directory:" -ForegroundColor DarkCyan
		Write-Host $operationLogOutput -ForegroundColor Green
		Write-Host "`n"
		
		exit
	}
}

##################
# GET Operations #
##################

if ($operationType -eq "GET") {
	# Use AD to generate separate files, if available
	# Otherwise, fail-over to command line call
	if ($psADModPresent) {
		# Get our user objects from AD, only get required properties
		# Pipe output into select and perform first file export of account attributes
		$users = Get-ADUser -SearchBase $searchBase -Filter * -Properties SamAccountName, mail, mailNickname, proxyAddresses |
		  select SamAccountName, mail, mailNickname |
		  Export-Csv $attributesOutput -NoType

		$operationLog += "Attributes file generated for " + $($users.Count) + " users`r`n`r`n"
		Write-Host "Attributes file generated for $($users.Count) users`n"

		$operationLog += "Beginning proxy address file generation`r`n"
		Write-Host "Beginning proxy address file generation`n"

		# We need to go through all user objects and remove any X400 addresses from the proxy addresses array
		foreach ($u in $users) {
			# Make an array list we can remove items from
			[System.Collections.ArrayList]$userProxies = $u.proxyAddresses
			
			# Create an empty array to hold any found X400 addresses
			$toRemove = @()
			foreach ($p in $userProxies) {
				if ($p -Match "X400") {
					$toRemove += $p
				}
			}
			
			# If any X400 addresses were found, remove them from the array list
			if ($toRemove.Count -gt 0) {
				foreach ($addr in $toRemove) {
					$userProxies.Remove($addr) 
				}
				
				# Write the modified array list back into the user object
				$u.proxyAddresses = $userProxies
			}
		}

		# Sort users by proxy addresses count for each individual
		# Use max proxy count to decide how many columns to generate in Proxy Address output file
		$maxProxy = $users | %{$_.proxyaddresses.count} | Sort-Object | Select-Object -Last 1
		
		# Loop through users and append each row to proxy address output file
		foreach ($u in $users) {
			$userProxies = [ordered]@{}
			$userProxies.Add("User",$u.SamAccountName)
			for ($i=0; $i -le $maxProxy; $i++) {
				if ($u.proxyaddresses[$I] -NotMatch "X400") {
					$userProxies.add("proxyaddress_$i",$u.proxyaddresses[$I])
				}
			}
			
			# Append each user's proxies into the output file
			[pscustomobject]$proxyAddress | Export-Csv $proxiesOutput -NoType -Append -Force
			Remove-Variable -Name proxyAddress
		}
		
		# Capture current date-time for logging
		$endDateTime = Get-Date -format "yyyy-MM-dd @ hh:mm"
		
		$operationLog += "Proxy Address File Generation Complete`r`n`r`n"
		$operationLog += "Exchange Migration Prep Script Finished at " + $endDateTime + "`r`n"
		
		$operationLogFile = [System.IO.StreamWriter] $operationLogOutput
		$operationLogFile.WriteLine($operationLog)
		$operationLogFile.close()
		
		Write-Host "Proxy Address File Generation Complete`n"
		Write-Host "Exchange Migration Prep Script Finished at $endDateTime`n"
		Write-Host "Check the output directory ($pathToWorkDir) for exported files and operation logs.`n"
	} else {
		# Handle GET operations via Command Line instead of AD Modules
	}
}

##################
# PUT Operations #
##################

if ($operationType -eq "PUT") {
	# Use Command Line statement to generate Exchange Migration data in a single compact file
	# TODO: Parse results into separate files for seamless failover
	
}
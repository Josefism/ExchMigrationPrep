# ExchMigrationPrep

A PowerShell script for exporting Exchange account attributes and proxy addresses from Active Directory accounts prior to migrating accounts to O365

## Summary
ExchMigrationPrep.ps1 automates the process of gathering MS Exchange account attributes (SamAccountName, Email Address, Name, Proxy Addresses) from Active Directory accounts prior to removal of Exchange from the on-premise environment during migration to Office 365. 
	
ExchMigrationPrep.ps1 also allows users to put exported data back into Active Directory following migration.

Both GET and PUT operations save exported data into the local filesystem at a user-specified location. All script operations are logged to local logfiles in the script's working directory.

### Environmental Assumptions:
* Windows Server environment running Active Directory
* Powershell >= v3.0 
* Active Directory Web Services running on at least one server in the environment for best results (failover option is built-in)

## Usage

First put ExchMigrationPrep.ps1 into a folder on your server. You will need to access the script locally within PowerShell to initiate execution.
	
Open a PowerShell command window and change directory to the folder containing the script. Run the script with or without parameters. The script will prompt for any required parameters that are not provided.

Respond to prompts for any required parameter corrections or additional information. If the script cannot access Active Directory PowerShell modules for GET operations, it will failover to legacy commands that produce a single export file.

GET operations using the Active Directory modules produce two export files. ExchAttributes export files contain Name, sAMAccountName, and primary Email. ExchProxies export files contain sAMAccountName and enumerated proxy addresses for each account.

GET operations using the legacy commands produce one export file containing all fields for each account. Proxy addresses for each account are contained in a single field for each row.

PUT operations require the Active Directory modules (which should be available in any O365-based environment). PUT operations target the files exported from GET operations in the working directory specified by the user.

All operations generate separate log files within the working directory.

## GET Parameters

-pathToWorkDir | -path
* Accepts directory path with drive letter
* Enter path with or without enclosing quotes
* Forward slashes will be auto-corrected to backslashes and trailing backslash will be added if needed
* Do not include filename
* If the drive letter and path format are valid, the path will be automatically created if it does not exist

-operationType | -operation
* Accepts GET or PUT
* Defaults to GET if the parameter contains unrecognized text

-searchBase | -search
* Accepts the full Distinguished Name of the top level Organizational Unit to search for user accounts within Active Directory
* Case-Insensitive
* Separate segments with commas (NO spaces BEFORE or AFTER commas)

___

## Credits, Comments, etc.

This script was built for automating parts of the migration process during client migrations from on-premise environments to O365. I'm hosting on GitHub for version control and bug tracking. Hopefully it is useful for others.

## TODO

* Improve error handling logic for some steps (i.e., Working Directory not found/cannot be created, auto-correction for SearchBase parameter formatting, etc.)
* Add PUT operation steps from old individual scripts
* Add PUT file selection prompting in case of multiple script runs in a single environment
* Add System/Application logging instead of (or in addition to) current script logs
* Automate legacy command failover data parsing to seamlessly produce the same file formats as AD Module exports
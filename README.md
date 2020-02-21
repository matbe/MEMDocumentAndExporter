# Microsoft Endpoint Manager ( aka Intune ) Documentation And Export tool

This repository contains PowerShell scripts to Document and Export settings from a Microsoft Endpoint Manager Tenant.

## They mayor benefits with this solution is the following:
- [x] Option to create a documentation of your tenant settings in a Word document.
- [x] Option to Export the settings in a json format that can be used as backup and in many cases used for import again.
- [x] Or BOTH!
- [x] Highly customizable.
- [x] Documents all assignment for every configuration, PolicySet and more.
- [x] Option to document who did the last change for every object.

The script creates a document with a nice looking Table of content for an easy overview:
<p align="center">
<img width="615" height="581" src="https://user-images.githubusercontent.com/15101419/74410682-ef1d5200-4e39-11ea-9897-df083e54dda4.png">
</p>

An example how a configuration item could look when documented:
<p align="center">
<img width="632" height="226" src="https://user-images.githubusercontent.com/15101419/74410285-0c9dec00-4e39-11ea-9110-bbb564f0a7ee.png">
</p>

### At the moment the following classes are supported:
* Managed Device Overview (Document only)
* Terms And Conditions
* Device Compliance Policies
* Device Enrollment Configurations
*  Device Configurations
* Windows Autopilot Deployment Profiles
* Mobile Apps
* Apple Push Notification Certificate (untested)
* VPP Tokens (untested)
* Policy Sets
* Group Policy Configurations
* Device Management Scripts (also exports the scripts)
* Groups

## Prerequisites
For the script to work the following 2 modules must be installed on the system
* AzureAD
* PSWriteWord

The account used when connecting to the tenant must have access to all objects to be documented/exported.

If the option *\<DocumentLastChange\>* is set to True in the configuration file the user must also have permission to read the audit logs in the tenant.

There is **NO** requirement that Word is installed on the system.

The script has only been tested with PowerShell 5.1.


# How to use it
The solution consists of 3 parts.
* The PowerShell script ***Export-MEMConfiguration.ps1***
* A xml file used for controlling the script ***Export-MEMConfiguration.xml***
* A word document used as a template ***MEMDocumentationTempl.docx***

### Export-MEMConfiguration.ps1
The script has a few parameters that can be used. The two most important are:
* `-Config` = Using this parameter you can specify a path and name to a custom configuration file. This can be useful if you have different tenants or settings you are using often.
* `-Force` = When running the script with the force setting it will not ask for confirmation when creating a folder when exporting settings.

These parameters are also available but the preferred way is to set them in the config file, if used the will override any settings in the config file:
* `-Tenant` = Specifies which tenant to connect to.
* `-ExportPath` = Path where the exported data will be created.
* `-DocumentName` = Name of the final Document.

### Export-MEMConfiguration.xml
The configuration xml consists of two parts. A Configuration part and a Process part. The Configuration part controls how the script operates and the Process part controls which object and classes should be documented and/or exported.
#### Configuration
`Tenant` = Specifies which tenant to use, using script parameter `-Tenant` will override this value.

`Document (True\False) ` = Enables the document feature and creates a word file.

`DocumentName` = Sets the document name, using script parameter `-DocumentName` will override this value.

`DocumentLastChange (True\False)` = Retrieves who did the last change for each object using audit logs where applicable. This requires the appropriate rights to read audit logs to work properly. 

`MaxStringLength (default 300)` = Sets the max length for any name/value when documenting. Values longer then specified value will be truncated.

`Export (True\False)` = Enables the export function, this will export a JSON file for each setting.

`ExportPath` = Sets the path where the exported files will be created, if files already there they will be overwritten, using script parameter `-ExportPath` will override this value. Using the value `{0}` in the string will replace it with current date and time (yyyyMMddHHmm). `C:\Temp\Export_{0}` would create a folder called `C:\Temp\Export_202002122300`. If both config and parameter are empty, script will ask for input. 

`ExportCSV (True\False)` = This will export a csv for every exported file in addition to the JSON created

`AppendDate (True\False)` = Appends the date to all exported files

`Maxlogfilesize (Default 5)` = Max log size in Megabytes

`graphApiVersion (Default Beta)` = API version, valid values are "Beta" and "v1.0". ***Must be Beta for the script to work***, might change in the future!
#### Process
Used for determine what to document and/or export. All values must be `True` or `False`.

### MEMDocumentationTempl.docx 
<p align="center">
<img width="606" height="252" src="https://user-images.githubusercontent.com/15101419/74409841-e297fa00-4e37-11ea-8e27-777d4b3f4ae3.png">
</p>

Word document created in OpenXML format used as a template.
The following values will be updated dynamically by the script:
* #TENANT#
* #DATE#
* #USERNAME#

## Running the script.
When the script is started it will first ask you to very that you are trying to connect to the correct tenant:

`Trying to connect to <yourtenant>.onmicrosoft.com, do you want to continue? Y or N?`

Then it will ask for the account that you wish to use:

`Please specify your user principal name for Azure Authentication`

Next it will ask you to confirm that you want to create the export folder (If the folder doesn’t already exist):

`Path 'C:\Temp\Export_202002091903' doesn't exist, do you want to create this directory? Y or N?`

From here on the script will document and/or export all classes that were enabled in the config file. The export will create a subfolder for each class to make it more readable.

<p align="center">
<img width="267" height="192" src="https://user-images.githubusercontent.com/15101419/74410064-7a95e380-4e38-11ea-80b7-c060a5b2c1b4.png">
</p>

Once finished it will try and open the word document that it has created.

## Credits
* [AzureAD](https://www.powershellgallery.com/packages/azuread)
* [PSWriteWord](https://github.com/EvotecIT/PSWriteWord)
* [The Graph documentation](https://docs.microsoft.com/en-us/graph/api/overview?toc=.%2Fref%2Ftoc.json&view=graph-rest-beta)
* [Microsoft Graph PowerShell Examples on Github](https://github.com/microsoftgraph/powershell-intune-samples)

## Bugs and feature requests
If you find any bugs or have an idea for a feature that is missing please create an issue.
For updates follow me here or on Twitter [matbg](https://twitter.com/matbg)

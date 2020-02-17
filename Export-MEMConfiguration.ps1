<#
.SYNOPSIS
  Export-MEMConfiguration

.DESCRIPTION
  Exports and documents all device configurations from a specified tenant.

.PARAMETER tenant
  Specifies which tentant to use. <sometenant.onmicrosoft.com>, overrides setting in the config file.

.PARAMETER ExportPath
  Path to export configuration. Will be created if not existing, overrides setting in the config file.

.PARAMETER DocumentName
  Set the Document name, overrides setting in the config file.

.PARAMETER Config
  Specifies the path to a custom config file, will override the default one.

.PARAMETER Force
  Skip confirmation to create folder if not existing.

.NOTES
  Version:        1.0
  Author:         Mattias Benninge
  Creation Date:  2020-01-07
  Purpose/Change: Initial script development

.EXAMPLE
    Export-MEMConfiguration.ps1

#>
#Requires -Modules AzureAD,PSWriteWord
#region --------------------------------------------------[Script Parameters]------------------------------------------------------
Param (
    [Parameter(Mandatory = $False)] [string]$Tenant = "",
    [Parameter(Mandatory = $False)] [string]$ExportPath = "",
    [Parameter(Mandatory = $False)] [string]$DocumentName = "",
    [Parameter(Mandatory = $False)] [string]$Config = "",
    [Parameter(Mandatory = $False)] [switch]$Force = $false
)
#endregion --------------------------------------------------[Script Parameters]------------------------------------------------------
#region ---------------------------------------------------[Declarations]----------------------------------------------------------
$DateTimeRegex = "\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{7}Z|\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z"
$script:User = ""

# Load Settings from Export-MEMConfiguration.xml
$global:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogPath = "$($global:ScriptPath)\$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log"
If ($null -eq $Config -or $Config -eq "") { $Config = Join-Path ($global:ScriptPath) "Export-MEMConfiguration.xml" }
$DocumentTemplate = Join-Path ($global:ScriptPath) "MEMDocumentationTempl.docx"
#endregion ---------------------------------------------------[Declarations]----------------------------------------------------------
#region ---------------------------------------------------[Functions]------------------------------------------------------------

#region Logging: Functions used for Logging, do not edit!
Function Start-Log {
    [CmdletBinding()]
    param (
        [ValidateScript( { Split-Path $_ -Parent | Test-Path })]
        [string]$FilePath
    )
	
    try {
        if (!(Test-Path $FilePath)) {
            ## Create the log file
            New-Item $FilePath -Type File | Out-Null
        }
		
        ## Set the global variable to be used as the FilePath for all subsequent Write-Log
        ## calls in this session
        $global:ScriptLogFilePath = $FilePath
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

Function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
		
        [Parameter()]
        [ValidateSet(1, 2, 3)]
        [int]$LogLevel = 1
    )    
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    
    if ($MyInvocation.ScriptName) {
        $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
    }
    else {
        #if the script havn't been saved yet and does not have a name this will state unknown.
        $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "Unknown", $LogLevel
    }
    $Line = $Line -f $LineFormat

    #Make sure the logfile do not exceed the $maxlogfilesize
    if (Test-Path $ScriptLogFilePath) { 
        if ((Get-Item $ScriptLogFilePath).length -ge $maxlogfilesize) {
            If (Test-Path "$($ScriptLogFilePath.Substring(0,$ScriptLogFilePath.Length-1))_") {
                Remove-Item -path "$($ScriptLogFilePath.Substring(0,$ScriptLogFilePath.Length-1))_" -Force
            }
            Rename-Item -Path $ScriptLogFilePath -NewName "$($ScriptLogFilePath.Substring(0,$ScriptLogFilePath.Length-1))_" -Force
        }
    }

    Add-Content -Value $Line -Path $ScriptLogFilePath

}
#endregion Logging: Functions used for Logging, do not edit!

# Add functions Here
function Get-AuthToken {
    <#
    .SYNOPSIS
    This function is used to authenticate with the Graph API REST interface
    .DESCRIPTION
    The function authenticate with the Graph API Interface with the tenant name
    .EXAMPLE
    Get-AuthToken
    Authenticates you with the Graph API interface
    .NOTES
    NAME: Get-AuthToken
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $User
    )
    
    Write-Host "Checking for AzureAD module..."
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    if ($null -eq $AadModule) {
        Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
        Write-Log "AzureAD PowerShell module not found, looking for AzureADPreview" -LogLevel 2
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    }
    
    if ($null -eq $AadModule) {
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        write-Log "AzureAD Powershell module not installed..." -LogLevel 3
        write-Log "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -LogLevel 2
        write-Log "Script can't continue..." -LogLevel 3
        exit 1
    }
    
    # Getting path to ActiveDirectory Assemblies
    # If the module count is greater than 1 find the latest version
    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
        $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
        # Checking if there are multiple versions of the same module found
        if ($AadModule.count -gt 1) {
            $aadModule = $AadModule | Select-Object -Unique
        }
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
    else {
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    $clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$script:Tenant"
    
    try {
        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
        # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
        # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
        $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
        $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
        $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters, $userId).Result
        # If the accesstoken is valid then create the authentication header
        if ($authResult.AccessToken) {
            # Creating header for Authorization token
            $authHeader = @{
                'Content-Type'  = 'application/json'
                'Authorization' = "Bearer " + $authResult.AccessToken
                'ExpiresOn'     = $authResult.ExpiresOn
            }
            return $authHeader
        }
        else {
            Write-Host
            Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
            Write-Host
            Write-Log "Authorization Access Token is null, please re-run authentication..." -LogLevel 3
            break
        }
    }
    catch {
        write-host $_.Exception.Message -f Red
        write-host $_.Exception.ItemName -f Red
        write-host
        Write-Log  $_.Exception.Message -LogLevel 3
        Write-Log  $_.Exception.ItemName -LogLevel 3
        break
    }
}

Function Update-AuthToken() {
    # Checking if authToken exists before running authentication
    if ($global:authToken) {
        # Setting DateTime to Universal time to work in all timezones
        $DateTime = (Get-Date).ToUniversalTime()
        # If the authToken exists checking when it expires
        $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes
        if ($TokenExpires -le 0) {
            write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
            Write-Log "Authentication Token expired $($TokenExpires.ToString()) minutes ago" -LogLevel 2
            write-host
            # Defining User Principal Name if not present
            if ($null -eq $script:User -or $script:User -eq "") {
                $script:User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
                Write-Log "Connecting using user: $($script:User)"
                Write-Host
            }
            Write-Log "Updating the authToken for the Graph API"
            $global:authToken = Get-AuthToken -User $User
        }
    }
    # Authentication doesn't exist, calling Get-AuthToken function
    else {
        if ($null -eq $script:User -or $script:User -eq "") {
            $script:User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            Write-Log "Connecting using user: $($script:User)"
            Write-Host
        }
        # Getting the authorization token
        Write-Log "Updating the authToken for the Graph API"
        $global:authToken = Get-AuthToken -User $script:User
    }
}

Function Get-GraphUri() {
    <#
    .SYNOPSIS
    This function is used to get a class or item from the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and gets any obects returned
    .EXAMPLE
    Get-GraphUri -ApiVersion "Beta" -Class "deviceManagement/deviceConfigurations/" -Value
    .NOTES
    NAME: Get-GraphUri
    #>
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $True)] [ValidateSet("Beta", "v1.0")] [string]$ApiVersion,
        [Parameter(Mandatory = $True)] [string]$Class,
        [Parameter(Mandatory = $False)] [string]$Id = "",
        [Parameter(Mandatory = $False)] [string]$OData = "",
        [Parameter(Mandatory = $False)] [switch]$Value
    )        
    $baseuri = "https://graph.microsoft.com/"
    
    $uri = $baseuri + $ApiVersion + "/" + $Class

    If ($Id -ne "") {
        If ($uri -notmatch '.+?\/$') { $uri = $uri + "/" + $id }
        else { $uri = $uri + $id }
    }

    If ($OData -ne "") { $uri = $uri + $OData }
    Write-Verbose "Connecting to $uri"
    Write-Log "Connecting to $uri"
    try {
        If ($Value) {
            $response = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
        }
        else {
            $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get
        }
    }
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Log "Response content:`n$responseBody" -LogLevel 3
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        Write-Log "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -LogLevel 3
    }
    return $response
}

function Invoke-GraphClass() {
    <#
    .SYNOPSIS
    This function is used to build the document and export data for the specified class in a reusable way
    .DESCRIPTION
    This function is used to build the document and export data for the specified class in a reusable way
    .EXAMPLE
    Invoke-GraphClass -Class "deviceManagement/deviceEnrollmentConfigurations"

    .NOTES
    NAME: Invoke-GraphClass
    #>
    param (
        [Parameter(Mandatory = $True)][string]$Class,
        [Parameter(Mandatory = $True)][array]$Title,
        [Parameter(Mandatory = $False)][array]$Properties,
        [Parameter(Mandatory = $False)][string]$PropForFileName = "",
        [Parameter(Mandatory = $False)][switch]$Value,
        [Parameter(Mandatory = $False)][switch]$GetLastChange
    )

    Update-AuthToken

    If ($Value) { [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value }
    else { [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class }

    If ($Document) { Add-WordText -WordDocument $WordDocument -Text $Title -HeadingType Heading1 -Supress $True }

    If($responsarray.Count -gt 0)
    {
        foreach ($response in $responsarray) {
            $classpath = $class -replace "/", "\"
            If ($Export) { 
                If ($PropForFileName -eq "") {
                    $JSONFileName = Export-JSONData -JSON $response -ExportPath "$ExportPath\$classpath" -Force 
                }
                else {
                    $JSONFileName = Export-JSONData -JSON $response -ExportPath "$ExportPath\$classpath" -FileName (Format-DataToString -Data $response.$PropForFileName)  -Force 
                } 
            }

            If ($Document) {    
                If ($Properties.Count -ne 0) {
                    $subvalues = $response | Select-Object -Property $Properties #|Select-Object -Property displayName,id,lastModifiedDateTime,description
                }
                else { $subvalues = $response }

                $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
                foreach ($prop in $subvalues.psobject.properties) {
                    $hashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
                }
        
                If($GetLastChange)
                {
                    $auditresponse = $null
                    $auditresponse = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/auditEvents" -OData "?`$filter=resources/any(d:d/resourceId eq '$($response.id)')&`$top=1" -Value
                    If($null -ne $auditresponse){
                        $hashtable["Last Change By"] = (Format-DataToString $($auditresponse.actor.userPrincipalName))
                        $hashtable["Last Change Action"] = (Format-DataToString $($auditresponse.activityOperationType))
                    }
                }

                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                If ($PropForFileName -eq "") {
                    Add-WordText -WordDocument $WordDocument -Text $response.displayName -HeadingType Heading3 -Supress $True
                }
                else {
                    Add-WordText -WordDocument $WordDocument -Text (Format-DataToString -Data $response.$PropForFileName) -HeadingType Heading3 -Supress $True
                }
                

                Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window  -Supress $True
                If ($export) { 
                    Add-WordText -WordDocument $WordDocument -Text 'Exported file:' -Supress $True
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$JSONFileName" -UrlLink "$ExportPath\$classpath\$JSONFileName" -Supress $True 
                }
            }
        }
    }
}

function Invoke-GraphClassExpand() {
    <#
    .SYNOPSIS
    This function is used to build the document and export data for the specified class in a reusable way. 
    .DESCRIPTION
    This function is used to build the document and export data for the specified class in a reusable way.
    This function enumerates assignments and maps them to the group. This is only usefull for documentation.
    .EXAMPLE
    Invoke-GraphClassExpand -Class "deviceManagement/deviceEnrollmentConfigurations"

    .NOTES
    NAME: Invoke-GraphClassExpand
    #>
    param (
        [Parameter(Mandatory = $True)][string]$Class,
        [Parameter(Mandatory = $True)][array]$Title,
        [Parameter(Mandatory = $False)][array]$Properties,
        [Parameter(Mandatory = $False)][string]$PropForFileName = "",
        [Parameter(Mandatory = $False)][switch]$Value,
        [Parameter(Mandatory = $False)][switch]$GetLastChange
    )

    Update-AuthToken

    If ($Value) { [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value }
    else { [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class }

    If ($Document) { Add-WordText -WordDocument $WordDocument -Text $Title -HeadingType Heading1 -Supress $True }
    If($responsarray.Count -gt 0)
    {
        foreach ($response in $responsarray) {
            $classpath = $class -replace "/", "\"
            If ($Export) { 
                If ($PropForFileName -eq "") {
                    $JSONFileName = Export-JSONData -JSON $response -ExportPath "$ExportPath\$classpath" -Force 
                }
                else {
                    $JSONFileName = Export-JSONData -JSON $response -ExportPath "$ExportPath\$classpath" -FileName (Format-DataToString -Data $response.$PropForFileName)  -Force 
                } 
            }

            If ($Document) {    
                If ($Properties.Count -ne 0) {
                    $subvalues = $response | Select-Object -Property $Properties #|Select-Object -Property displayName,id,lastModifiedDateTime,description
                }
                else { $subvalues = $response }

                $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
                foreach ($prop in $subvalues.psobject.properties) {
                    $hashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
                }

                If($GetLastChange)
                {
                    $auditresponse = $null
                    $auditresponse = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/auditEvents" -OData "?`$filter=resources/any(d:d/resourceId eq '$($response.id)')&`$top=1" -Value
                    If($null -ne $auditresponse){
                        $hashtable["Last Change By"] = (Format-DataToString $($auditresponse.actor.userPrincipalName))
                        $hashtable["Last Change Action"] = (Format-DataToString $($auditresponse.activityOperationType))
                    }
                }

                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                If ($PropForFileName -eq "") {
                    Add-WordText -WordDocument $WordDocument -Text $response.displayName -HeadingType Heading3 -Supress $True
                }
                else {
                    Add-WordText -WordDocument $WordDocument -Text (Format-DataToString -Data $response.$PropForFileName) -HeadingType Heading3 -Supress $True
                }
                

                Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window  -Supress $True
                If ($export) { 
                    Add-WordText -WordDocument $WordDocument -Text 'Exported file:' -Supress $True
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$JSONFileName" -UrlLink "$ExportPath\$classpath\$JSONFileName" -Supress $True 
                }
                $expandeditem = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $response.id -OData '?$Expand=assignments'
                If (($expandeditem.assignments).Count -ge 1) {
                    Add-WordText -WordDocument $WordDocument -Text 'This item have been assigned to the following groups' -Supress $True
                    $ListOfGroups = @()
                    foreach ($assignment in $expandeditem.assignments) {
                        $ListOfGroups += ($Groups -match $assignment.target.groupId).displayName
                    }
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfGroups -Supress $True -Verbose
                }
            }
        }
    }
}

Function Export-JSONData() {
    <#
    .SYNOPSIS
    This function is used to export JSON data returned from Graph
    .DESCRIPTION
    This function is used to export JSON data returned from Graph
    .EXAMPLE
    Export-JSONData -JSON $JSON
    Export the JSON inputted on the function
    .NOTES
    NAME: Export-JSONData
    #>
    param (
        [Parameter(Mandatory = $True)]$JSON,
        [Parameter(Mandatory = $True)][string]$ExportPath,
        [Parameter(Mandatory = $False)][string]$FileName = "",
        [Parameter(Mandatory = $False)][switch]$Force
    )
    try {
        if ($JSON -eq "" -or $null -eq $JSON) {
            write-host "No JSON specified, please specify valid JSON..." -f Red
        }
        elseif (!$ExportPath) {
            write-host "No export path parameter set, please provide a path to export the file" -f Red
        }
        elseif (!(Test-Path $ExportPath)) {
            If (!$Force) {
                # If the directory path doesn't exist prompt user to create the directory
                Write-Host "Path '$ExportPath' doesn't exist, do you want to create this directory? Y or N?" -ForegroundColor Yellow
                $Confirm = read-host
                if ($Confirm -eq "y" -or $Confirm -eq "Y") {
                    new-item -ItemType Directory -Path "$ExportPath" | Out-Null
                    Write-Host
                }
                else {
                    Write-Host "Creation of directory path was cancelled, can't export JSON Data" -ForegroundColor Red
                    Write-Host
                    break
                }
            }    
            else {
                if (!(Test-Path "$ExportPath")) {
                    new-item -ItemType Directory -Path "$ExportPath" | Out-Null
                }
            }
        }

        $JSON1 = ConvertTo-Json $JSON -Depth 5
        $JSON_Convert = $JSON1 | ConvertFrom-Json
        If ($FileName -eq "") {
            $displayName = $JSON_Convert.displayName
        }
        else {
            $displayName = $FileName
        }
        # Updating display name to follow file naming conventions - https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx
        $DisplayName = $DisplayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"
        $Properties = ($JSON_Convert | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" }).Name
        If ($script:AppendDate) {
            $FileName_JSON = "$DisplayName" + "_" + $(get-date -f dd-MM-yyyy-H-mm-ss) + ".json"
        }
        else {
            $FileName_JSON = "$DisplayName" + ".json"
        }

        #write-verbose "Export Path: $ExportPath"
        
        If ($ExportCSV) {
            If ($script:AppendDate) {
                $FileName_CSV = "$DisplayName" + "_" + $(get-date -f dd-MM-yyyy-H-mm-ss) + ".csv"
            } 
            else {
                $FileName_CSV = "$DisplayName" + ".csv"
            }
            $Object = New-Object System.Object

            foreach ($Property in $Properties) {
                $Object | Add-Member -MemberType NoteProperty -Name $Property -Value $JSON_Convert.$Property
            }
    
            $Object | Export-Csv -LiteralPath "$ExportPath\$FileName_CSV" -Delimiter "," -NoTypeInformation -Append -Encoding UTF8
            write-host "CSV created in $ExportPath\$FileName_CSV..." -f cyan
        }
        
        $JSON1 | Set-Content -LiteralPath "$ExportPath\$FileName_JSON" -Encoding UTF8
        write-host "JSON created in $ExportPath\$FileName_JSON..." -f cyan
        write-log "JSON created in $ExportPath\$FileName_JSON..."

        return $FileName_JSON
    }
    catch {
        $_.Exception
    }
}

Function Format-DataToString() {
    <#
    .SYNOPSIS
    This function formats data returned from Graph
    .DESCRIPTION
    This function formats data returned from Graph
    .EXAMPLE
    Format-DataToString -Data $data
    .NOTES
    NAME: Format-DataToString
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $false)][AllowEmptyString()][AllowNull()] $Data
    )
    if ($Data -is [array]) {
        $Data = $Data -join ","
    }
    $Data = $Data -replace '^@{|}|^#microsoft.graph.|^@odata.', ""
    if ($Data -match $DateTimeRegex) {
        try {
            [DateTime]$Date = ([DateTime]::Parse($Data))
            $Data = "$($Date.ToShortDateString()) $($Date.ToShortTimeString())"
        }
        catch {
        }
    }
    If ($Data.Length -ge $MaxStringLength) {
        $Data = $Data.substring(0, $MaxStringLength) + "..."
    }
    return $Data
}
#endregion ---------------------------------------------------[Functions]------------------------------------------------------------

#-----------------------------------------------------------[Execution]------------------------------------------------------------
#Default logging to %temp%\scriptname.log, change if needed.
Start-Log -FilePath $LogPath
Write-Log "---------- Script Starting ----------"
# Load config.xml
if (Test-Path $Config) {
    try { 
        $Xml = [xml](Get-Content -Path $Config -Encoding UTF8)
        Write-Log -Message "Successfully loaded $Config" 
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Log -Message "Error, could not read $Config" -LogLevel 3
        Write-Log -Message "Error message: $ErrorMessage" -LogLevel 3
        Exit 1
    }
}
else {
    Write-Log -Message "Error, could not find or access $Config" -LogLevel 3
    Exit 1
}

# Load xml content into variables
try {
    #Load Configuration
    [bool]$Document = [System.Convert]::ToBoolean($Xml.root.Configuration.Document)
    [bool]$Export = [System.Convert]::ToBoolean($Xml.root.Configuration.Export)
    if ($DocumentName -eq "") { [string]$DocumentName = [string]$Xml.root.Configuration.DocumentName }
    [bool]$DocumentLastChange = [System.Convert]::ToBoolean($Xml.root.Configuration.DocumentLastChange)
    $MaxStringLength = [int]$Xml.root.Configuration.MaxStringLength
    $maxlogfilesize = [int]($Xml.root.Configuration.maxlogfilesize) * 1Mb
    $script:graphApiVersion = [string]$Xml.root.Configuration.graphApiVersion
    [bool]$script:AppendDate = [System.Convert]::ToBoolean($Xml.root.Configuration.AppendDate)
    if ($Tenant -eq "") { [string]$Tenant = [string]$Xml.root.Configuration.Tenant }
    if ($ExportPath -eq "") { [string]$ExportPath = ([string]$Xml.root.Configuration.ExportPath) -f (Get-Date -Format "yyyyMMddHHmm") }
    [bool]$ExportCSV = [System.Convert]::ToBoolean($Xml.root.Configuration.ExportCSV)

    #Load Process
    [bool]$ProcessmanagedDeviceOverview = [System.Convert]::ToBoolean($xml.root.Process.managedDeviceOverview)
    [bool]$ProcesstermsAndConditions = [System.Convert]::ToBoolean($xml.root.Process.termsAndConditions)
    [bool]$ProcessdeviceCompliancePolicies = [System.Convert]::ToBoolean($xml.root.Process.deviceCompliancePolicies)
    [bool]$ProcessdeviceEnrollmentConfigurations = [System.Convert]::ToBoolean($xml.root.Process.deviceEnrollmentConfigurations)
    [bool]$ProcessdeviceConfigurations = [System.Convert]::ToBoolean($xml.root.Process.deviceConfigurations)
    [bool]$ProcesswindowsAutopilotDeploymentProfiles = [System.Convert]::ToBoolean($xml.root.Process.windowsAutopilotDeploymentProfiles)
    [bool]$ProcessmobileApps = [System.Convert]::ToBoolean($xml.root.Process.mobileApps)
    [bool]$ProcessapplePushNotificationCertificate = [System.Convert]::ToBoolean($xml.root.Process.applePushNotificationCertificate)
    [bool]$ProcessvppTokens = [System.Convert]::ToBoolean($xml.root.Process.vppTokens)
    [bool]$Processpolicysets = [System.Convert]::ToBoolean($xml.root.Process.policysets)
    [bool]$ProcessgroupPolicyConfigurations = [System.Convert]::ToBoolean($xml.root.Process.groupPolicyConfigurations)
    [bool]$ProcessdeviceManagementScripts = [System.Convert]::ToBoolean($xml.root.Process.deviceManagementScripts)
    [bool]$ProcessGroups = [System.Convert]::ToBoolean($xml.root.Process.Groups)
    Write-Log -Message "Successfully processed all settings in $Config"
}
catch {
    Write-Log -Message "Xml content from $Config was not loaded properly" -LogLevel 3
    Exit 1
}

$script:Tenant = $Tenant

#region Start
# Make sure that at least one of $export or $document is set to "True", otherwise its pointless to run the script...
If ($export -eq $false -and $Document -eq $false) {
    Write-Error "At least one of the Document or Export variables must be set to True."
    Write-Log "At least one of the Document or Export variables must be set to True." -LogLevel 3
    Exit 1
}

Write-Host "Trying to connect to $script:Tenant, do you want to continue? Y or N?" -ForegroundColor Yellow
        
$Confirm = read-host
if ($Confirm -eq "y" -or $Confirm -eq "Y") {
    Write-Host "Connecting to $script:Tenant.."
    Write-Log "Connecting to $script:Tenant.."
}
else {
    Write-Host "Aborting..." -ForegroundColor Red
    Write-Log "Scipt was user aborted when connecting to tennant." -LogLevel 3
    exit 1
}

Update-AuthToken

Write-Host "Connected using $script:User.."
Write-Log "Connected using $script:User.."

# If Export = True, verify that exportpath is set and can be created.

if ($ExportPath -eq "") {
    $ExportPath = Read-Host -Prompt "Please specify a path to export the policy data and/or the documentation to e.g. C:\MEMOutput"
}

$ExportPath = $ExportPath.replace('"', '')
If (!$Force) {
    # If the directory path doesn't exist prompt user to create the directory
    if (!(Test-Path "$ExportPath")) {
        Write-Host
        Write-Host "Path '$ExportPath' doesn't exist, do you want to create this directory? Y or N?" -ForegroundColor Yellow
        $Confirm = read-host
        if ($Confirm -eq "y" -or $Confirm -eq "Y") {
            new-item -ItemType Directory -Path "$ExportPath" | Out-Null
        }
        else {
            Write-Host "Creation of directory path was cancelled..." -ForegroundColor Red
            Write-Log "Creation of directory path was cancelled..." -LogLevel 3
            exit 1
        }
    }
}    
else {
    if (!(Test-Path "$ExportPath")) {
        new-item -ItemType Directory -Path "$ExportPath" | Out-Null
    }
}



$FullDocumentationPath = Join-Path ($ExportPath) $DocumentName
Write-Log "Document Template = $DocumentTemplate"
Write-Log "Export path set to: $ExportPath"
Write-Log "Full Documentation Path = $FullDocumentationPath"

####################################################
# Set up Documentation
If ($Document) { 
    $WordDocument = Get-WordDocument -FilePath $DocumentTemplate

    foreach ($Paragraph in $WordDocument.Paragraphs) {
        $Paragraph.ReplaceText('#DATE#', (Get-Date -Format "yyyy.MM.dd HH:mm"))
        $Paragraph.ReplaceText('#TENANT#', $script:Tenant)
        $Paragraph.ReplaceText('#USERNAME#', $script:User)
    }

    Add-WordPageBreak -WordDocument $WordDocument -Supress $True
    Add-WordTOC -WordDocument $WordDocument -Title 'Table of content' -HeaderStyle Heading1 -Supress $True
    Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
}

#endregion Start
#Get all groups, used for resolving assignments where needed.
$groups = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "groups" -Value

#region managedDeviceOverview
If ($ProcessmanagedDeviceOverview -and $Document) { #No point in exporting this class
    Update-AuthToken

    $class = "deviceManagement/managedDeviceOverview"
    $DeviceOverview = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class
    $DeviceTable = $DeviceOverview | Select-Object -Property enrolledDeviceCount,mdmEnrolledCount,dualEnrolledDeviceCount,managedDeviceModelsAndManufacturers,lastModifiedDateTime
   
    $DThashtable = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($prop in $DeviceTable.psobject.properties) {
        $DThashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
    }

    $OShashtable = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($prop in $DeviceOverview.deviceOperatingSystemSummary.psobject.properties) {
        $OShashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
    }

    $EAhashtable = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($prop in $DeviceOverview.deviceExchangeAccessStateSummary.psobject.properties) {
        $EAhashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
    }
    
    Add-WordText -WordDocument $WordDocument -Text 'Device Overview' -HeadingType Heading1 -Supress $True
    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
    Add-WordTable -WordDocument $WordDocument -DataTable $DThashtable -Design LightGridAccent1 -AutoFit Window -Supress $True
    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
    Add-WordBarChart -WordDocument $WordDocument -ChartName 'Operating System Summary' -Names $OShashtable.Keys -Values $OShashtable.Values -NoLegend -BarDirection Bar
    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
    Add-WordBarChart -WordDocument $WordDocument -ChartName 'Exchange Access State Summary' -Names $EAhashtable.Keys -Values $EAhashtable.Values -NoLegend -BarDirection Bar
    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

    $DThashtable = $null
    $OShashtable = $null
    $EAhashtable = $null
}
#endregion managedDeviceOverview

# Get all classes that dont need extra treatment.
If ($ProcesstermsAndConditions) { Invoke-GraphClass -Class "deviceManagement/termsAndConditions" -Title 'Terms and Conditions' -PropForFileName "@odata.type" -Value }
If ($ProcessdeviceCompliancePolicies) { Invoke-GraphClassExpand -Class "deviceManagement/deviceCompliancePolicies" -Title 'Device Compliance Policies' -Value -GetLastChange:$DocumentLastChange}
If ($ProcessdeviceEnrollmentConfigurations) { Invoke-GraphClass -Class "deviceManagement/deviceEnrollmentConfigurations" -Title 'Device Enrollment Configurations' -PropForFileName "@odata.type" -Value -GetLastChange:$DocumentLastChange}
If ($ProcessdeviceConfigurations) { Invoke-GraphClassExpand -Class "deviceManagement/deviceConfigurations" -Title 'Device Configurations' -Properties "displayName", "id", "lastModifiedDateTime", "description" -Value -GetLastChange:$DocumentLastChange}
If ($ProcesswindowsAutopilotDeploymentProfiles) { Invoke-GraphClassExpand -Class "deviceManagement/windowsAutopilotDeploymentProfiles" -Title 'Windows Autopilot Deployment Profiles' -Value -GetLastChange:$DocumentLastChange}
If ($ProcessmobileApps) { Invoke-GraphClassExpand -Class "deviceAppManagement/mobileApps" -Title 'Mobile Apps' -Properties "displayName", "id", "lastModifiedDateTime", "description" -Value -GetLastChange:$DocumentLastChange}
If ($ProcessapplePushNotificationCertificate) { Invoke-GraphClass -Class "deviceManagement/applePushNotificationCertificate" -Value -Title 'Apple Push Notification Certificate' }
If ($ProcessvppTokens) { Invoke-GraphClassExpand -Class "deviceAppManagement/vppTokens" -Title 'VPP Tokens' -Value }

#region policySets
If ($Processpolicysets) {
    Update-AuthToken

    $class = "deviceAppManagement/policysets"
    [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value

    If ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Policy Sets' -HeadingType Heading1 -Supress $True }
    If($responsarray.Count -gt 0)
    {
        foreach ($response in $responsarray) {
            $classpath = $class -replace "/", "\"
            $expandeditem = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $response.id -OData '?$Expand=assignments,items'
            If ($Export) { $JSONFileName = Export-JSONData -JSON $expandeditem -ExportPath "$ExportPath\$classpath" -Force }

            If ($Document) {    

                $subvalues = $response 

                $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
                foreach ($prop in $subvalues.psobject.properties) {
                    $hashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
                }
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-WordText -WordDocument $WordDocument -Text $response.displayName -HeadingType Heading3 -Supress $True
                Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window  -Supress $True
                If ($export) { 
                    Add-WordText -WordDocument $WordDocument -Text 'Exported files:' -Supress $True
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$JSONFileName" -UrlLink "$ExportPath\$classpath\$JSONFileName" -Supress $True 
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$($response.filename)" -UrlLink "$ExportPath\$classpath\$($response.filename)" -Supress $True
                }
            
                If (($expandeditem.items).Count -ge 1) {
                    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                    Add-WordText -WordDocument $WordDocument -Text 'This sets includes the following items' -Supress $True
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $expandeditem.items.displayname -Supress $True -Verbose
                }

                If (($expandeditem.assignments).Count -ge 1) {
                    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                    Add-WordText -WordDocument $WordDocument -Text 'This item have been assigned to the following groups' -Supress $True
                    $ListOfGroups = @()
                    foreach ($assignment in $expandeditem.assignments) {
                        $ListOfGroups += ($Groups -match $assignment.target.groupId).displayName
                    }
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfGroups -Supress $True -Verbose
                }
            }
        }
    }
    $responsarray = $null
}
#endregion policySets
#region groupPolicyConfigurations
If ($ProcessgroupPolicyConfigurations) {
    Update-AuthToken

    $class = "deviceManagement/groupPolicyConfigurations"
    [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value
    $classpath = $class -replace "/", "\"
    If ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Group Policy Configurations' -HeadingType Heading1 -Supress $True }
    If($responsarray.Count -gt 0)
    {
        foreach ($response in $responsarray) {
            
            $expandeditem = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $response.id -OData '?$Expand=assignments'
            
            $pvJSONFileNames = @()
            $gpcJSONFileNames = @()

            If ($Export) { $JSONFileName = Export-JSONData -JSON $expandeditem -ExportPath "$ExportPath\$classpath" -Force }
            
            $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
            
            $dvclass = "deviceManagement/groupPolicyConfigurations/$($response.id)/definitionValues"
            $gpcarr = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $dvclass -Value
            
            If ($Document) { Add-WordText -WordDocument $WordDocument -Text $response.displayName -HeadingType Heading3 -Supress $True }

            $i = 1
            foreach ($gpc in $gpcarr) {
                If ($Export) { $gpcJSONFileNames += Export-JSONData -JSON $gpc -ExportPath "$ExportPath\$classpath\definitionValues" -Filename "$($response.displayName)_dv_$i" -Force }

                $gpcclass = "deviceManagement/groupPolicyConfigurations/$($response.id)/definitionValues/$($gpc.id)/presentationValues"
                $pvarr = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $gpcclass -Value
                $j = 1    
                Foreach ($pv in $pvarr) {
                    If ($Export) { $pvJSONFileNames += Export-JSONData -JSON $pv -ExportPath "$ExportPath\$classpath\definitionValues\presentationValues" -Filename "$($response.displayName)_pv_$j" -Force }
                    
                    $pvclass = "deviceManagement/groupPolicyConfigurations/$($response.id)/definitionValues/$($gpc.id)/presentationValues/$($pv.id)/presentation"
                    $presentation = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $pvclass
                    
                    $value = $null
                    If (!$null -eq $pv.values) {
                        $value = $pv.values
                    }
                    elseif (!$null -eq $pv.value) {
                        $value = $pv.value
                    }
                    $hashtable[(Format-DataToString $($presentation.label))] = (Format-DataToString $($value))
                    $j++
                }
                $i++
            }
            If ($Document) {
                Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window  -Supress $True
                If ($export) { 
                    Add-WordText -WordDocument $WordDocument -Text 'Exported files:' -Supress $True
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$JSONFileName" -UrlLink "$ExportPath\$classpath\$JSONFileName" -Supress $True
                    If (($gpcJSONFileNames).Count -ge 1) { Add-WordHyperLink -WordDocument $WordDocument -UrlText "$($response.displayName)_definitionValues" -UrlLink "$ExportPath\$classpath\definitionValues\" -Supress $True } 
                    If (($pvJSONFileNames).Count -ge 1) { Add-WordHyperLink -WordDocument $WordDocument -UrlText "$($response.displayName)_presentationValues" -UrlLink "$ExportPath\$classpath\definitionValues\presentationValues\" -Supress $True } 
                }

                If (($expandeditem.assignments).Count -ge 1) {
                    Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                    Add-WordText -WordDocument $WordDocument -Text 'This item have been assigned to the following groups' -Supress $True
                    $ListOfGroups = @()
                    foreach ($assignment in $expandeditem.assignments) {
                        $ListOfGroups += ($Groups -match $assignment.target.groupId).displayName
                    }
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfGroups -Supress $True -Verbose
                }
            }
        }
    }
    $responsarray = $null
}
#endregion groupPolicyConfigurations
#region exportscripts
If ($ProcessdeviceManagementScripts) {
    Update-AuthToken

    $class = "deviceManagement/deviceManagementScripts"
    [array]$responsarray = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value

    If ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Device Management Scripts' -HeadingType Heading1 -Supress $True }
    If($responsarray.Count -gt 0)
    {
        foreach ($response in $responsarray) {
            $classpath = $class -replace "/", "\"
            $expandeditem = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $response.id -OData '?$Expand=assignments'
            If ($Export) { 
                $JSONFileName = Export-JSONData -JSON $response -ExportPath "$ExportPath\$classpath" -Force 
                [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($expandeditem.scriptContent)) | Out-File -FilePath $(Join-Path -Path "$ExportPath\$classpath" -ChildPath "$($response.filename)")
            }

            If ($Document) {    

                $subvalues = $response | select-object -ExcludeProperty scriptContent

                $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
                foreach ($prop in $subvalues.psobject.properties) {
                    $hashtable[(Format-DataToString $($prop.Name))] = (Format-DataToString $($prop.Value))
                }
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-WordText -WordDocument $WordDocument -Text $response.displayName -HeadingType Heading3 -Supress $True
                Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window -Supress $True
                If ($export) { 
                    Add-WordText -WordDocument $WordDocument -Text 'Exported files:' -Supress $True
                    Add-WordHyperLink -WordDocument $WordDocument -UrlText "$JSONFileName" -UrlLink "$ExportPath\$classpath\$JSONFileName" -Supress $True 
                }
            
                If (($expandeditem.assignments).Count -ge 1) {
                    Add-WordText -WordDocument $WordDocument -Text 'This item have been assigned to the following groups' -Supress $True
                    $ListOfGroups = @()
                    foreach ($assignment in $expandeditem.assignments) {
                        $ListOfGroups += ($Groups -match $assignment.target.groupId).displayName
                    }
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfGroups -Supress $True -Verbose
                }
            }
        }
    }
    $responsarray = $null
}
#endregion exportscripts
#region groups
If ($ProcessGroups) {
    If ($Export) {
        foreach ($group in $Groups) {
            #write-host "Group:$($group.displayName) $($group.id)" -f Yellow
            $JSONFileName = Export-JSONData -JSON $group -ExportPath "$ExportPath\Groups" -Force
            #Write-Host
        }
    }
    If ($Document) {
        $hashtable = New-Object System.Collections.Specialized.OrderedDictionary
        $hashtable = $groups | Sort-Object -Property displayName | Select-Object -Property displayName, @{name = 'groupTypes'; expression = { $_.groupTypes -join "," } }, renewedDateTime
    
        Add-WordText -WordDocument $WordDocument -Text 'Groups' -HeadingType Heading1 -Supress $True 
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
        Add-WordTable -WordDocument $WordDocument -DataTable $hashtable -Design LightGridAccent1 -AutoFit Window  -Supress $True
    }
}
#endregion exportgroups

If ($Document) { Save-WordDocument -WordDocument $WordDocument -FilePath $FullDocumentationPath -OpenDocument -Supress $true }

Write-Log "---------- Script Completed ----------"



<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 24.06.14.0926

param (
    [ValidateSet("Online", "Onprem")][Parameter(Mandatory = $false)]
    [string]$Environment="Online",

    [Parameter(Position=0, Mandatory=$True, HelpMessage="The Mailbox parameter specifies the mailbox to be accessed.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$false, HelpMessage="The PermissionType parameter specifies whether the app registrations uses delegated or application permissions")] [ValidateSet('Application','Delegated')]
    [string]$PermissionType,
    
    [Parameter(Mandatory=$False, HelpMessage="The Archive parameter is a switch to search the archive mailbox (otherwise, the main mailbox is searched).")]
    [alias("SearchArchive")] [switch]$Archive,

    [Parameter(Mandatory=$False, HelpMessage="The ProcessSubfolders parameter is a switch to enable searching the subfolders of any specified folder.")]
    [switch]$ProcessSubfolders,

    [Parameter(Mandatory=$False, HelpMessage="The IncludeFolderList parameter specifies the folder(s) to be searched (if not present, then the Inbox folder will be searched).  Any exclusions override this list.")]
    $IncludeFolderList,

    [Parameter(Mandatory=$False, HelpMessage="The ExcludeFolderList parameter specifies the folder(s) to be excluded (these folders will not be searched).")]
    $ExcludeFolderList,

    [Parameter(Mandatory=$false, HelpMessage="The SearchDumpster parameter is a switch to search the recoverable items.")]
    [switch]$SearchDumpster,

    [Parameter(Mandatory=$False, HelpMessage="The MessageClass parameter specifies the message class of the items being searched.")]
    [ValidateNotNullOrEmpty()]
    [string]$MessageClass,

    [Parameter(Mandatory=$false, HelpMessage="The CreatedBefore parameter specifies only messages created before this date will be searched.")]
    [DateTime]$CreatedBefore,

    [Parameter(Mandatory=$false, HelpMessage="The CreatedAfter parameter specifies only messages created after this date will be searched.")]
    [DateTime]$CreatedAfter,

    [Parameter(Mandatory=$False, HelpMessage="The Subject parameter specifies the subject string used by the search.")]
    [string]$Subject,

    [Parameter(Mandatory=$False, HelpMessage="The Sender parameter specifies the sender email address used by the search.")]
    [string]$Sender,

    [Parameter(Mandatory=$False, HelpMessage="The MessageBody parameter specifies the body string used by the search.")]
    [string]$MessageBody,

    [Parameter(Mandatory=$False, HelpMessage="The MessageId parameter specified the MessageId used by the search.")]
    [string]$MessageId,

    [ValidateSet("Global", "USGovernmentL4", "USGovernmentL5", "ChinaCloud")]
    [Parameter(Mandatory = $false)]
    [string]$AzureEnvironment = "Global",

    [Parameter(Mandatory=$False, HelpMessage="The EwsUrl parameter specifies the on-premises Exchange Web Services URL.")]
    [string]$EwsUrl,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 2147483)]
    [int]$TimeoutSeconds = 300,

    [ValidateScript({ Test-Path $_ })]
    [Parameter(Mandatory = $false)]
    [string]$DLLPath,

    [Parameter(Mandatory=$False, HelpMessage="The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.")]
    [string]$OAuthClientId,

    [Parameter(Mandatory=$False, HelpMessage="The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).")]
    [string]$OAuthTenantId,

    [Parameter(Mandatory=$False, HelpMessage="The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.")]
    [string]$OAuthRedirectUri = "http://localhost:8004",

    [Parameter(Mandatory=$False, HelpMessage="The OAuthSecretKey parameter is the the secret for the registered application.")]
    [SecureString]$OAuthClientSecret,

    [Parameter(Mandatory=$False, HelpMessage="The OAuthCertificate parameter is the certificate for the registered application. Certificate auth requires MSAL libraries to be available.")]
    $OAuthCertificate = $null,

    [Parameter(Mandatory=$False,HelpMessage="The CertificateStore parameter specifies the certificate store where the certificate is loaded.")] [ValidateSet("CurrentUser", "LocalMachine")]
     [string] $CertificateStore = $null,
     
    [Parameter(Mandatory=$false)] [Array]$Scope= @("EWS.AccessAsUser.All"),

    [Parameter(Mandatory=$false)][switch]$UseImpersonation,

    [Parameter(Mandatory=$false)][pscredential]$credential,

    [ValidateScript({ Test-Path $_ })] [Parameter(Mandatory = $true, HelpMessage="The OutputPath parameter specifies the path for the EWS usage report.")] [string] $OutputPath,

    [Parameter(Mandatory=$False, HelpMessage="The ThrottlingDelay parameter specifies the throttling delay (time paused between sending EWS requests) - note that this will be increased automatically if throttling is detected")]
    [int]$ThrottlingDelay = 0,

    [Parameter(Mandatory=$False, HelpMessage="The BatchSize parameter specifies the batch size (number of items batched into one EWS request) - this will be decreased if throttling is detected")]
    [int]$BatchSize = 200
)
begin {
function Write-VerboseLog ($Message) {
    $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
}

function Write-HostLog ($Message) {
    $Script:Logger = $Script:Logger | Write-LoggerInstance $Message
}

function Enable-TrustAnyCertificateCallback {
    param()

    <#
        This helper function can be used to ignore certificate errors. It works by setting the ServerCertificateValidationCallback
        to a callback that always returns true. This is useful when you are using self-signed certificates or certificates that are
        not trusted by the system.
    #>

    Add-Type -TypeDefinition @"
    namespace Microsoft.CSSExchange {
        public class CertificateValidator {
            public static bool TrustAnyCertificateCallback(
                object sender,
                System.Security.Cryptography.X509Certificates.X509Certificate cert,
                System.Security.Cryptography.X509Certificates.X509Chain chain,
                System.Net.Security.SslPolicyErrors sslPolicyErrors) {
                return true;
            }

            public static void IgnoreCertificateErrors() {
                System.Net.ServicePointManager.ServerCertificateValidationCallback = TrustAnyCertificateCallback;
            }
        }
    }
"@
    [Microsoft.CSSExchange.CertificateValidator]::IgnoreCertificateErrors()
}

function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    process {
        $consoleHost = $host.Name -eq "ConsoleHost"

        if ($null -ne $Script:WriteHostManipulateObjectAction) {
            $Object = & $Script:WriteHostManipulateObjectAction $Object
        }

        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params

        if ($null -ne $Script:WriteHostDebugAction -and
            $null -ne $Object) {
            &$Script:WriteHostDebugAction $Object
        }
    }
}

function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

function SetWriteHostAction ($DebugAction) {
    $Script:WriteHostDebugAction = $DebugAction
}

function SetWriteHostManipulateObjectAction ($ManipulateObject) {
    $Script:WriteHostManipulateObjectAction = $ManipulateObject
}

function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {

        if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
            $Message = & $Script:WriteVerboseManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Verbose $Message

        if ($null -ne $Script:WriteVerboseDebugAction) {
            & $Script:WriteVerboseDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteVerboseDebugAction) {
            & $Script:WriteRemoteVerboseDebugAction $Message
        }
    }
}

function SetWriteVerboseAction ($DebugAction) {
    $Script:WriteVerboseDebugAction = $DebugAction
}

function SetWriteRemoteVerboseAction ($DebugAction) {
    $Script:WriteRemoteVerboseDebugAction = $DebugAction
}

function SetWriteVerboseManipulateMessageAction ($DebugAction) {
    $Script:WriteVerboseManipulateMessageAction = $DebugAction
}

function Write-Warning {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Warning from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )
    process {

        if ($null -ne $Script:WriteWarningManipulateMessageAction) {
            $Message = & $Script:WriteWarningManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Warning $Message

        # Add WARNING to beginning of the message by default.
        $Message = "WARNING: $Message"

        if ($null -ne $Script:WriteWarningDebugAction) {
            & $Script:WriteWarningDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteWarningDebugAction) {
            & $Script:WriteRemoteWarningDebugAction $Message
        }
    }
}

function SetWriteWarningAction ($DebugAction) {
    $Script:WriteWarningDebugAction = $DebugAction
}

function SetWriteRemoteWarningAction ($DebugAction) {
    $Script:WriteRemoteWarningDebugAction = $DebugAction
}

function SetWriteWarningManipulateMessageAction ($DebugAction) {
    $Script:WriteWarningManipulateMessageAction = $DebugAction
}

function Get-NewLoggerInstance {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = (Get-Location).Path,

        [ValidateNotNullOrEmpty()]
        [string]$LogName = "Script_Logging",

        [bool]$AppendDateTime = $true,

        [bool]$AppendDateTimeToFileName = $true,

        [int]$MaxFileSizeMB = 10,

        [int]$CheckSizeIntervalMinutes = 10,

        [int]$NumberOfLogsToKeep = 10
    )

    $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
    $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)

    if (-not (Test-Path $LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
        } catch {
            throw "Failed to create Log Directory: $LogDirectory. Inner Exception: $_"
        }
    }

    return [PSCustomObject]@{
        FullPath                 = $fullFilePath
        AppendDateTime           = $AppendDateTime
        MaxFileSizeMB            = $MaxFileSizeMB
        CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
        NumberOfLogsToKeep       = $NumberOfLogsToKeep
        BaseInstanceFileName     = $fileName.Replace(".txt", "")
        Instance                 = 1
        NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
        PreventLogCleanup        = $false
        LoggerDisabled           = $false
    } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
}

function Write-LoggerInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]$Object
    )
    process {
        if ($LoggerInstance.LoggerDisabled) { return }

        if ($LoggerInstance.AppendDateTime -and
            $Object.GetType().Name -eq "string") {
            $Object = "[$([System.DateTime]::Now)] : $Object"
        }

        # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
        $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false

        #Upkeep of the logger information
        if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }

        #Set next update time to avoid issues so we can log things
        $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
        $item = Get-ChildItem $LoggerInstance.FullPath

        if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
            $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
            $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
            $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
            $LoggerInstance.Instance++
            $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)

            $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"

            if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                $item | Remove-Item -Force
            }
        }
    }
    end {
        return $LoggerInstance
    }
}

function Invoke-LoggerInstanceCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance
    )
    process {
        if ($LoggerInstance.LoggerDisabled -or
            $LoggerInstance.PreventLogCleanup) {
            return
        }

        Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
            Remove-Item -Force
    }
}

function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [ScriptBlock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

function Test-ADCredentials {
    [CmdletBinding()]
    [OutputType([System.Object])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    <#
        This function tests whether the credentials provided are valid by trying to connect to LDAP server using Kerberos authentication.
        It returns a PSCustomObject with two properties:
        - UsernameFormat: "local", "upn" or "downlevel" depending on the format of the username provided
        - CredentialsValid: $true if the credentials are valid, $false if they are not valid, $null if the function was unable to perform the validation
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $credentialsValid = $null
        # Username formats: https://learn.microsoft.com/windows/win32/secauthn/user-name-formats
        $usernameFormat = "local"
        try {
            Add-Type -AssemblyName System.DirectoryServices.Protocols -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to load System.DirectoryServices.Protocols"
            Write-Verbose "Exception: $_"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    process {
        $domain = $Credentials.GetNetworkCredential().Domain
        if ([System.String]::IsNullOrEmpty($domain)) {
            Write-Verbose "Domain is empty which could be an indicator that UPN was passed instead of domain\username"
            $domain = ($Credentials.GetNetworkCredential().UserName).Split("@")
            if ($domain.Count -eq 2) {
                Write-Verbose "Domain was extracted from UPN"
                $domain = $domain[-1]
                $usernameFormat = "upn"
            } else {
                Write-Verbose "Failed to extract domain from UPN - seems that username was passed without domain and so cannot be validated"
                $domain = $null
            }
        } else {
            Write-Verbose "Username was provided in down-level logon name format"
            $usernameFormat = "downlevel"
        }

        if (-not([System.String]::IsNullOrEmpty($domain))) {
            $ldapDirectoryIdentifier = New-Object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($domain)
            # Use Kerberos authentication as NTLM might lead to false/positive results in case the password was changed recently
            $ldapConnection = New-Object -TypeName System.DirectoryServices.Protocols.LdapConnection($ldapDirectoryIdentifier, $Credentials, [DirectoryServices.Protocols.AuthType]::Kerberos)
            # Enable Kerberos encryption (sign and seal)
            $ldapConnection.SessionOptions.Signing = $true
            $ldapConnection.SessionOptions.Sealing = $true
            try {
                $ldapConnection.Bind()
                Write-Verbose "Connection succeeded with credentials"
                $credentialsValid = $true
            } catch [System.DirectoryServices.Protocols.LdapException] {
                if ($_.Exception.ErrorCode -eq 49) {
                    # ErrorCode 49 means invalid credentials
                    Write-Verbose "Failed to connect to LDAP server with credentials provided"
                    $credentialsValid = $false
                } else {
                    Write-Verbose "Failed to connect to LDAP server for other reason"
                    Write-Verbose "ErrorCode: $($_.Exception.ErrorCode)"
                }
                Write-Verbose "Exception: $_"
                Invoke-CatchActionError $CatchActionFunction
            } catch {
                Write-Verbose "Exception occurred while connecting to LDAP server - unable to perform credential validation"
                Write-Verbose "Exception: $_"
                Invoke-CatchActionError $CatchActionFunction
            }
        }
    }
    end {
        if ($null -ne $ldapConnection) {
            $ldapConnection.Dispose()
        }
        return [PSCustomObject]@{
            UsernameFormat   = $usernameFormat
            CredentialsValid = $credentialsValid
        }
    }
}

function Show-Disclaimer {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [ValidateNotNullOrEmpty()]
        [string]$Target,
        [ValidateNotNullOrEmpty()]
        [string]$Operation
    )

    if ($PSCmdlet.ShouldProcess($Message, $Target, $Operation) -or
        $WhatIfPreference) {
        return
    } else {
        exit
    }
}

function EWSAuth {
    param(
        [string]$Environment,
        $Token,
        $EwsUri
    )
    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    ## Create the Exchange Service object with credentials
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    $Service.Timeout = $TimeoutSeconds * 1000
    if ($Environment -eq "Onprem") {
        if([string]::IsNullOrEmpty($OAuthClientId)){
            $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($credential.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)))
        }
        else {
            $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token)
        }
    }
    else {
        $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token)
    }

    $Service.Url = $EwsUri
    $Service.HttpHeaders.Add("X-AnchorMailbox", $Mailbox)
    if($PermissionType -eq "Application" -or $UseImpersonation) {
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox)
    }
    
    return $Service
}

function Get-CloudServiceEndpoint {
    [CmdletBinding()]
    param(
        [string]$EndpointName
    )

    <#
        This shared function is used to get the endpoints for the Azure and Microsoft 365 services.
        It returns a PSCustomObject with the following properties:
            GraphApiEndpoint: The endpoint for the Microsoft Graph API
            ExchangeOnlineEndpoint: The endpoint for Exchange Online
            AutoDiscoverSecureName: The endpoint for Autodiscover
            AzureADEndpoint: The endpoint for Azure Active Directory
            EnvironmentName: The name of the Azure environment
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
    }
    process {
        # https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        switch ($EndpointName) {
            "Global" {
                $environmentName = "AzureCloud"
                $graphApiEndpoint = "https://graph.microsoft.com"
                $exchangeOnlineEndpoint = "https://outlook.office.com"
                $autodiscoverSecureName = "https://autodiscover-s.outlook.com"
                $azureADEndpoint = "https://login.microsoftonline.com"
                break
            }
            "USGovernmentL4" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                break
            }
            "USGovernmentL5" {
                $environmentName = "AzureUSGovernment"
                $graphApiEndpoint = "https://dod-graph.microsoft.us"
                $exchangeOnlineEndpoint = "https://outlook-dod.office365.us"
                $autodiscoverSecureName = "https://autodiscover-s-dod.office365.us"
                $azureADEndpoint = "https://login.microsoftonline.us"
                break
            }
            "ChinaCloud" {
                $environmentName = "AzureChinaCloud"
                $graphApiEndpoint = "https://microsoftgraph.chinacloudapi.cn"
                $exchangeOnlineEndpoint = "https://partner.outlook.cn"
                $autodiscoverSecureName = "https://autodiscover-s.partner.outlook.cn"
                $azureADEndpoint = "https://login.partner.microsoftonline.cn"
                break
            }
        }
    }
    end {
        return [PSCustomObject]@{
            EnvironmentName        = $environmentName
            GraphApiEndpoint       = $graphApiEndpoint
            ExchangeOnlineEndpoint = $exchangeOnlineEndpoint
            AutoDiscoverSecureName = $autodiscoverSecureName
            AzureADEndpoint        = $azureADEndpoint
        }
    }
}

function Get-NewJsonWebToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$CertificateThumbprint,
        [ValidateSet("CurrentUser", "LocalMachine")][Parameter(Mandatory = $false)][string]$CertificateStore = "CurrentUser",
        [Parameter(Mandatory = $false)][string]$Issuer,
        [Parameter(Mandatory = $false)][string]$Audience,
        [Parameter(Mandatory = $false)][string]$Subject,
        [Parameter(Mandatory = $false)][int]$TokenLifetimeInSeconds = 3600,
        [ValidateSet("RS256", "RS384", "RS512")][Parameter(Mandatory = $false)][string]$SigningAlgorithm = "RS256"
    )
    
    <#
        Shared function to create a signed Json Web Token (JWT) by using a certificate.
        It is also possible to use a secret key to sign the token, but that is not supported in this function.
        The function returns the token as a string if successful, otherwise it returns $null.
        https://www.rfc-editor.org/rfc/rfc7519
        https://learn.microsoft.com/azure/active-directory/develop/active-directory-certificate-credentials
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
    }
    process {
        try {
            $certificate = Get-ChildItem Cert:\$CertificateStore\My\$CertificateThumbprint
            if ($certificate.HasPrivateKey) {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                # Base64url-encoded SHA-1 thumbprint of the X.509 certificate's DER encoding
                $x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
                $x5t = ((($x5t).Replace("\+", "-")).Replace("/", "_")).Replace("=", "")
                Write-Verbose "x5t is: $x5t"
            } else {
                Write-Verbose "We don't have a private key for certificate: $CertificateThumbprint and so cannot sign the token"
                return
            }
        } catch {
            Write-Verbose "Unable to import the certificate - Exception: $($Error[0].Exception.Message)"
            return
        }

        $header = [ordered]@{
            alg = $SigningAlgorithm
            typ = "JWT"
            x5t = $x5t
        }

        # "iat" (issued at) and "exp" (expiration time) must be UTC and in UNIX time format
        $payload = @{
            iat = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds)
            exp = [Math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date -Date "01/01/1970")).TotalSeconds) + $TokenLifetimeInSeconds
        }

        # Issuer, Audience and Subject are optional as per RFC 7519
        if (-not([System.String]::IsNullOrEmpty($Issuer))) {
            Write-Verbose "Issuer: $Issuer will be added to payload"
            $payload.Add("iss", $Issuer)
        }

        if (-not([System.String]::IsNullOrEmpty($Audience))) {
            Write-Verbose "Audience: $Audience will be added to payload"
            $payload.Add("aud", $Audience)
        }

        if (-not([System.String]::IsNullOrEmpty($Subject))) {
            Write-Verbose "Subject: $Subject will be added to payload"
            $payload.Add("sub", $Subject)
        }

        $headerJson = $header | ConvertTo-Json -Compress
        $payloadJson = $payload | ConvertTo-Json -Compress

        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($headerJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($payloadJson)).Split("=")[0].Replace("+", "-").Replace("/", "_")

        $signatureInput = [System.Text.Encoding]::ASCII.GetBytes("$headerBase64.$payloadBase64")

        Write-Verbose "Header (Base64) is: $headerBase64"
        Write-Verbose "Payload (Base64) is: $payloadBase64"
        Write-Verbose "Signature input is: $signatureInput"

        $signingAlgorithmToUse = switch ($SigningAlgorithm) {
            ("RS384") { [Security.Cryptography.HashAlgorithmName]::SHA384 }
            ("RS512") { [Security.Cryptography.HashAlgorithmName]::SHA512 }
            default { [Security.Cryptography.HashAlgorithmName]::SHA256 }
        }
        Write-Verbose "Signing the Json Web Token using: $SigningAlgorithm"

        $signature = $privateKey.SignData($signatureInput, $signingAlgorithmToUse, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signature = [Convert]::ToBase64String($signature).Split("=")[0].Replace("+", "-").Replace("/", "_")
    }
    end {
        if ((-not([System.String]::IsNullOrEmpty($headerBase64))) -and
            (-not([System.String]::IsNullOrEmpty($payloadBase64))) -and
            (-not([System.String]::IsNullOrEmpty($signature)))) {
            Write-Verbose "Returning Json Web Token"
            return ("$headerBase64.$payloadBase64.$signature")
        } else {
            Write-Verbose "Unable to create Json Web Token"
            return
        }
    }
}

function Get-ApplicationAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$TenantID,
        [Parameter(Mandatory = $true)][string]$ClientID,
        [Parameter(Mandatory = $true)][string]$Secret,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $false)][string]$TokenService = "oauth2/v2.0/token",
        [Parameter(Mandatory = $false)][switch]$CertificateBasedAuthentication,
        [Parameter(Mandatory = $true)][string]$Scope
    )
    
    <#
        Shared function to create an OAuth token by using a JWT or secret.
        If you want to use a certificate, set the CertificateBasedAuthentication switch and pass a JWT token as the Secret parameter.
        You can use the Get-NewJsonWebToken function to create a JWT token.
        If you want to use a secret, pass the secret as the Secret parameter.
        This function returns a PSCustomObject with the OAuth token, status and the time the token was created.
        If the request fails, the PSCustomObject will contain the exception message.
    #>
    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        $oAuthTokenCallSuccess = $false
        $exceptionMessage = $null

        Write-Verbose "TenantID: $TenantID - ClientID: $ClientID - Endpoint: $Endpoint - TokenService: $TokenService - Scope: $Scope"
        $body = @{
            scope      = $Scope
            client_id  = $ClientID
            grant_type = "client_credentials"
        }

        if ($CertificateBasedAuthentication) {
            Write-Verbose "Function was called with CertificateBasedAuthentication switch"
            $body.Add("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer")
            $body.Add("client_assertion", $Secret)
        } else {
            Write-Verbose "Authentication is based on a secret"
            $body.Add("client_secret", $Secret)
        }

        $invokeRestMethodParams = @{
            ContentType = "application/x-www-form-urlencoded"
            Method      = "POST"
            Body        = $body # Create string by joining bodyList with '&'
            Uri         = "$Endpoint/$TenantID/$TokenService"
        }
    }
    process {
        try {
            Write-Verbose "Now calling the Invoke-RestMethod cmdlet to create an OAuth token"
            $oAuthToken = Invoke-RestMethod @invokeRestMethodParams
            Write-Verbose "Invoke-RestMethod call was successful"
            $oAuthTokenCallSuccess = $true
        } catch {
            Write-Host "We fail to create an OAuth token - Exception: $($_.Exception.Message)" -ForegroundColor Red
            $exceptionMessage = $_.Exception.Message
        }
    }
    end {
        return [PSCustomObject]@{
            OAuthToken           = $oAuthToken
            Successful           = $oAuthTokenCallSuccess
            ExceptionMessage     = $exceptionMessage
            LastTokenRefreshTime = (Get-Date)
        }
    }
}

function CheckTokenExpiry {
    param(
            $ApplicationInfo,
            [ref]$EWSService,
            [ref]$Token,
            [string]$Environment,
            $EWSOnlineURL,
            $AuthScope,
            $AzureADEndpoint
        )
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
    # if token is going to expire in next 5 min then refresh it
    if ($null -eq $script:tokenLastRefreshTime -or $script:tokenLastRefreshTime.AddMinutes(55) -lt (Get-Date)) {
        Write-Verbose "Requesting new OAuth token as the current token expires at $($script:tokenLastRefreshTime)."
        if($PermissionType -eq "Application") {
        $createOAuthTokenParams = @{
            TenantID                       = $ApplicationInfo.TenantID
            ClientID                       = $ApplicationInfo.ClientID
            Endpoint                       = $AzureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($ApplicationInfo.CertificateThumbprint)))
            Scope                           = $Script:TokenScope
        }

        # Check if we use an app secret or certificate by using regex to match Json Web Token (JWT)
        if ($ApplicationInfo.AppSecret -match "^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)") {
            $jwtParams = @{
                CertificateThumbprint = $ApplicationInfo.CertificateThumbprint
                CertificateStore      = $CertificateStore
                Issuer                = $ApplicationInfo.ClientID
                Audience              = "$AzureADEndpoint/$($ApplicationInfo.TenantID)/oauth2/v2.0/token"
                Subject               = $ApplicationInfo.ClientID
            }
            $jwt = Get-NewJsonWebToken @jwtParams

            if ($null -eq $jwt) {
                Write-Host "Unable to sign a new Json Web Token by using certificate: $($ApplicationInfo.CertificateThumbprint)" -ForegroundColor Red
                exit
            }

            $createOAuthTokenParams.Add("Secret", $jwt)
        } else {
            $createOAuthTokenParams.Add("Secret", $ApplicationInfo.AppSecret)
        }

        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to refresh EWS OAuth token. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        Write-Host "Obtained a new token" -ForegroundColor Green
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        }
        else {
            # Request an authorization code from the Microsoft Azure Active Directory endpoint
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ApplicationInfo.ClientID
                    scope         = $AuthScope
                    grant_type    = "refresh_token"
                    refresh_token =  $Script:RefreshToken
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $script:tokenLastRefreshTime = (Get-Date)
                $Script:RefreshToken = $tokens.refresh_token
                $Script:Token = $tokens.access_token
            } 
            else {
                Write-Host "Unable to redeem the authorization code for an access token." -ForegroundColor Red
                exit
            }
        }
    }
}

function Get-DelegatedAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$AzureADEndpoint = "https://login.microsoftonline.com",
        [Parameter(Mandatory = $true)][string]$Scope,
        [Parameter(Mandatory = $false)][string]$ClientID,
        [Parameter(Mandatory = $false)][string]$RedirectUri
    )

    <#
        This function is used to get an access token for the Azure Graph API by using the OAuth 2.0 authorization code flow
        with PKCE (Proof Key for Code Exchange). The OAuth 2.0 authorization code grant type, or auth code flow,
        enables a client application to obtain authorized access to protected resources like web APIs.
        The auth code flow requires a user-agent that supports redirection from the authorization server
        (the Microsoft identity platform) back to your application.

        More information about the auth code flow with PKCE can be found here:
        https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow#protocol-details
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
       
        $responseType = "code" # Provides the code as a query string parameter on our redirect URI
        $prompt = "select_account" # We want to show the select account dialog
        $codeChallengeMethod = "S256" # The code challenge method is S256 (SHA256)
        $codeChallengeVerifier = Get-NewS256CodeChallengeVerifier
        $state = ([guid]::NewGuid()).Guid
        $connectionSuccessful = $false
    }
    process {
        $codeChallenge = $codeChallengeVerifier.CodeChallenge
        $codeVerifier = $codeChallengeVerifier.Verifier

        # Request an authorization code from the Microsoft Azure Active Directory endpoint
        $authCodeRequestUrl = "$AzureADEndpoint/organizations/oauth2/v2.0/authorize?client_id=$clientId" +
        "&response_type=$responseType&redirect_uri=$redirectUri&scope=$scope&state=$state&prompt=$prompt" +
        "&code_challenge_method=$codeChallengeMethod&code_challenge=$codeChallenge"

        Start-Process -FilePath $authCodeRequestUrl
        $authCodeResponse = Start-LocalListener

        if ($null -ne $authCodeResponse) {
            # Redeem the returned code for an access token
            $redeemAuthCodeParams = @{
                Uri             = "$AzureADEndpoint/organizations/oauth2/v2.0/token"
                Method          = "POST"
                ContentType     = "application/x-www-form-urlencoded"
                Body            = @{
                    client_id     = $ClientID
                    scope         = $Scope
                    code          = ($($authCodeResponse.Split("=")[1]).Split("&")[0])
                    redirect_uri  = $RedirectUri
                    grant_type    = "authorization_code"
                    code_verifier = $codeVerifier
                }
                UseBasicParsing = $true
            }
            $redeemAuthCodeResponse = Invoke-WebRequestWithProxyDetection -ParametersObject $redeemAuthCodeParams

            if ($redeemAuthCodeResponse.StatusCode -eq 200) {
                $tokens = $redeemAuthCodeResponse.Content | ConvertFrom-Json
                $connectionSuccessful = $true
            } else {
                Write-Host "Unable to redeem the authorization code for an access token." -ForegroundColor Red
            }
        } else {
            Write-Host "Unable to acquire an authorization code from the Microsoft Azure Active Directory endpoint." -ForegroundColor Red
        }
    }
    end {
        if ($connectionSuccessful) {
            return [PSCustomObject]@{
                AccessToken = $tokens.access_token
                RefreshToken = $tokens.refresh_token
                #TenantId    = (Convert-JsonWebTokenToObject $tokens.id_token).Payload.tid
                LastTokenRefreshTime = (Get-Date)
                Successful           = $true
            }
        }
        exit
    }
}

function Convert-JsonWebTokenToObject {
    param(
        [Parameter(Mandatory = $true)][ValidatePattern("^([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_=]+)\.([a-zA-Z0-9_\-\+\/=]*)")][string]$Token
    )

    <#
        This function can be used to split a JSON web token (JWT) into its header, payload, and signature.
        The JWT is expected to be in the format of <header>.<payload>.<signature>.
        The function returns a PSCustomObject with the following properties:
            Header    - The header of the JWT
            Payload   - The payload of the JWT
            Signature - The signature of the JWT

            It returns $null if the JWT is not in the expected format or conversion fails.
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        function ConvertJwtFromBase64StringWithoutPadding {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Jwt
            )
            $Jwt = ($Jwt.Replace("-", "+")).Replace("_", "/")
            switch ($Jwt.Length % 4) {
                0 { return [System.Convert]::FromBase64String($Jwt) }
                2 { return [System.Convert]::FromBase64String($Jwt + "==") }
                3 { return [System.Convert]::FromBase64String($Jwt + "=") }
                default { throw "The JWT is not a valid Base64 string." }
            }
        }
    }
    process {
        $tokenParts = $Token.Split(".")
        $tokenHeader = $tokenParts[0]
        $tokenPayload = $tokenParts[1]
        $tokenSignature = $tokenParts[2]

        Write-Verbose "Now processing token header..."
        $tokenHeaderDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenHeader))

        Write-Verbose "Now processing token payload..."
        $tokenPayloadDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenPayload))

        Write-Verbose "Now processing token signature..."
        $tokenSignatureDecoded = [System.Text.Encoding]::UTF8.GetString((ConvertJwtFromBase64StringWithoutPadding $tokenSignature))
    }
    end {
        if (($null -ne $tokenHeaderDecoded) -and
            ($null -ne $tokenPayloadDecoded) -and
            ($null -ne $tokenSignatureDecoded)) {
            Write-Verbose "Conversion of the token was successful"
            return [PSCustomObject]@{
                Header    = ($tokenHeaderDecoded | ConvertFrom-Json)
                Payload   = ($tokenPayloadDecoded | ConvertFrom-Json)
                Signature = $tokenSignatureDecoded
            }
        }

        Write-Verbose "Conversion of the token failed"
        return $null
    }
}

function Get-NewS256CodeChallengeVerifier {
    param()

    <#
        This function can be used to generate a new SHA256 code challenge and verifier following the PKCE specification.
        The Proof Key for Code Exchange (PKCE) extension describes a technique for public clients to mitigate the threat
        of having the authorization code intercepted. The technique involves the client first creating a secret,
        and then using that secret again when exchanging the authorization code for an access token.

        The function returns a PSCustomObject with the following properties:
        Verifier: The verifier that was generated
        CodeChallenge: The code challenge that was generated

        It returns $null if the code challenge and verifier generation fails.

        More information about the auth code flow with PKCE can be found here:
        https://www.rfc-editor.org/rfc/rfc7636
    #>

    Write-Verbose "Calling $($MyInvocation.MyCommand)"

    $bytes = [System.Byte[]]::new(64)
    ([System.Security.Cryptography.RandomNumberGenerator]::Create()).GetBytes($bytes)
    $b64String = [Convert]::ToBase64String($bytes)
    $verifier = (($b64String.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

    $newMemoryStream = [System.IO.MemoryStream]::new()
    $newStreamWriter = [System.IO.StreamWriter]::new($newMemoryStream)
    $newStreamWriter.write($verifier)
    $newStreamWriter.Flush()
    $newMemoryStream.Position = 0
    $hash = Get-FileHash -InputStream $newMemoryStream | Select-Object Hash
    $hex = $hash.Hash

    $bytesArray = [byte[]]::new($hex.Length / 2)

    for ($i = 0; $i -lt $hex.Length; $i+=2) {
        $bytesArray[$i/2] = [Convert]::ToByte($hex.Substring($i, 2), 16)
    }

    $base64Encoded = [Convert]::ToBase64String($bytesArray)
    $base64UrlEncoded = (($base64Encoded.TrimEnd("=")).Replace("+", "-")).Replace("/", "_")

    if ((-not([System.String]::IsNullOrEmpty($verifier))) -and
        (-not([System.String]::IsNullOrEmpty(($base64UrlEncoded))))) {
        Write-Verbose "Verifier and CodeChallenge generated successfully"
        return [PSCustomObject]@{
            Verifier      = $verifier
            CodeChallenge = $base64UrlEncoded
        }
    }

    Write-Verbose "Verifier and CodeChallenge generation failed"
    return $null
}

function Start-LocalListener {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only non-destructive operations are performed in this function.')]
    param(
        [Parameter(Mandatory = $false)][int]$Port = 8004,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 60,
        [Parameter(Mandatory = $false)][string]$UrlContains = "code=",
        [Parameter(Mandatory = $false)][string]$ExpectedHttpMethod = "GET",
        [Parameter(Mandatory = $false)][string]$ResponseOutput = "Authentication complete. You can return to the application. Feel free to close this browser tab."
    )

    <#
        This function is used to start a local listener on the specified port (default: 8004).
        It will wait for the specified amount of seconds (default: 60) for a request to be made.
        The function will return the URL of the request that was made.
    #>

    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand)"
        $url = $null
        $signalled = $false
        $stopwatch = New-Object System.Diagnostics.Stopwatch
        $listener = New-Object Net.HttpListener
    }
    process {
        $listener.Prefixes.add("http://localhost:$($Port)/")
        try {
            Write-Verbose "Starting listener..."
            Write-Verbose "Listening on port: $($Port)"
            Write-Verbose "Waiting $($TimeoutSeconds) seconds for request to be made to url that contains: $($UrlContains)"
            $stopwatch.Start()
            $listener.Start()

            while ($listener.IsListening) {
                $task = $listener.GetContextAsync()

                while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                    if ($task.AsyncWaitHandle.WaitOne(100)) {
                        $signalled = $true
                        break
                    }
                    Start-Sleep -Milliseconds 100
                }

                if ($signalled) {
                    $context = $task.GetAwaiter().GetResult()
                    $request = $context.Request
                    $response = $context.Response
                    $url = $request.RawUrl
                    $content = [byte[]]@()

                    if (($url.Contains($UrlContains)) -and
                        ($request.HttpMethod -eq $ExpectedHttpMethod)) {
                        Write-Verbose "Request made to listener and url that was called is as expected. HTTP Method: $($request.HttpMethod)"
                        $content = [System.Text.Encoding]::UTF8.GetBytes($ResponseOutput)
                        $response.StatusCode = 200 # OK
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    } else {
                        Write-Verbose "Request made to listener but the url that was called is not as expected. URL: $($url)"
                        $response.StatusCode = 404 # Not Found
                        $response.OutputStream.Write($content, 0, $content.Length)
                        $response.Close()
                        break
                    }
                } else {
                    Write-Verbose "Timeout of $($TimeoutSeconds) seconds reached..."
                    break
                }
            }
        } finally {
            Write-Verbose "Stopping listener..."
            Start-Sleep -Seconds 2
            $stopwatch.Stop()
            $listener.Stop()
        }
    }
    end {
        return $url
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")][string]$Uri,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][switch]$UseBasicParsing,
        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")][hashtable]$ParametersObject,
        [Parameter(Mandatory = $false, ParameterSetName = "Default")][string]$OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)][string]$TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")][string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Get-OAuthToken {
    param(
        [array]$AppScope,
        [Parameter(Mandatory=$true)] [ValidateSet('EWS','Graph')][string]$Api
    )
    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    $EwsUri = $Script:EwsEndpoint.Substring(0,$Script:EwsEndpoint.IndexOf("/EWS/"))
    if($PermissionType -eq "Application") {
        switch ($Api) {
            "EWS" {
                
                $Script:TokenScope = "$($EwsUri)/.default"
            }
            "Graph" {
                $Script:TokenScope = "$($cloudService.graphApiEndpoint)/.default"
            }
        }

        if ([System.String]::IsNullOrEmpty($OAuthCertificate)) {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($OAuthClientSecret)
            $Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Script:applicationInfo.Add("AppSecret", $Secret)
        }
        else {
            $jwtParams = @{
                CertificateThumbprint = $OAuthCertificate
                CertificateStore      = $CertificateStore
                Issuer                = $OAuthClientId
                Audience              = "$azureADEndpoint/$OAuthTenantId/oauth2/v2.0/token"
                Subject               = $OAuthClientId
            }
            $jwt = Get-NewJsonWebToken @jwtParams
    
            if ($null -eq $jwt) {
                Write-Host "Unable to generate Json Web Token by using certificate: $CertificateThumbprint" -ForegroundColor Red
                exit
            }
    
            $Script:applicationInfo.Add("AppSecret", $jwt)
            $Script:applicationInfo.Add("CertificateThumbprint", $OAuthCertificate)
        }
    
        $createOAuthTokenParams = @{
            TenantID                       = $OAuthTenantId
            ClientID                       = $OAuthClientId
            Secret                         = $Script:applicationInfo.AppSecret
            Scope                          = $Script:TokenScope
            Endpoint                       = $azureADEndpoint
            CertificateBasedAuthentication = (-not([System.String]::IsNullOrEmpty($OAuthCertificate)))
        }
    
        #Create OAUTH token
        $oAuthReturnObject = Get-ApplicationAccessToken @createOAuthTokenParams
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }
        $Script:Token = $oAuthReturnObject.OAuthToken.access_token
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
    }
    elseif ($PermissionType -eq "Delegated") {
        if(-not(($AppScope.Contains("email")))) {
            $AppScope += "email"
        }
        if(-not(($AppScope.Contains("openid")))) {
            $AppScope += "openid"
        }
        if(-not(($AppScope.Contains("offline_access")))) {
            $AppScope += "offline_access"
        }
        switch ($Api) {
            "EWS" {
                $Script:TokenScope = "$($EwsUri)//$($Scope)"
            }
            "Graph" {
                $Script:TokenScope = "$($cloudService.GraphApiEndpoint)//$($Scope)"
            }
        }
        
        $oAuthReturnObject = Get-DelegatedAccessToken -AzureADEndpoint $cloudService.AzureADEndpoint -Scope $Script:TokenScope -ClientID $OAuthClientId -RedirectUri $OAuthRedirectUri
        if ($oAuthReturnObject.Successful -eq $false) {
            Write-Host ""
            Write-Host "Unable to fetch an OAuth token for accessing EWS. Please review the error message below and re-run the script:" -ForegroundColor Red
            Write-Host $oAuthReturnObject.ExceptionMessage -ForegroundColor Red
            exit
        }    
        $Script:tokenLastRefreshTime = $oAuthReturnObject.LastTokenRefreshTime
        $Script:Token = $oAuthReturnObject.AccessToken
        $Script:RefreshToken = $oAuthReturnObject.RefreshToken
    }
}

function LoadEWSManagedAPI {
    $path = $DLLPath

    if ([System.String]::IsNullOrEmpty($path)) {
        Write-Host "Trying to find Microsoft.Exchange.WebServices.dll in the script folder"
        $path = (Get-ChildItem -LiteralPath "$PSScriptRoot\EWS" -Recurse -Filter "Microsoft.Exchange.WebServices.dll" -ErrorAction SilentlyContinue |
                Select-Object -First 1).FullName

        if ([System.String]::IsNullOrEmpty($path)) {
            Write-Host "Microsoft.Exchange.WebServices.dll wasn't found - attempting to download it from the internet" -ForegroundColor Yellow
            $nuGetPackage = Get-NuGetPackage -PackageId "Microsoft.Exchange.WebServices" -Author "Microsoft"

            if ($nuGetPackage.DownloadSuccessful) {
                $unzipNuGetPackage = Invoke-ExtractArchive -CompressedFilePath $nuGetPackage.NuGetPackageFullPath -TargetFolder "$PSScriptRoot\EWS"

                if ($unzipNuGetPackage.DecompressionSuccessful) {
                    $path = (Get-ChildItem -Path $unzipNuGetPackage.FullPathToDecompressedFiles -Recurse -Filter "Microsoft.Exchange.WebServices.dll" |
                            Select-Object -First 1).FullName
                } else {
                    Write-Host "Failed to unzip Microsoft.Exchange.WebServices.dll. Please unzip the package manually." -ForegroundColor Red
                    exit
                }
            } else {
                Write-Host "Failed to download Microsoft.Exchange.WebServices.dll from the internet. Please download the package manually and extract the dll. Provide the path to dll using DLLPath parameter." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "Microsoft.Exchange.WebServices.dll was found in the script folder" -ForegroundColor Green
        }
    }

    if ($path -notlike "*Microsoft.Exchange.WebServices.dll") {
        $path = "$path\Microsoft.Exchange.WebServices.dll"
    }

    try {
        Import-Module -Name $path -ErrorAction Stop
        return $true
    } catch {
        Write-Host "Failed to import Microsoft.Exchange.WebServices.dll Inner Exception`n`n$_" -ForegroundColor Red
        exit
    }
}
function CreateService($smtpAddress, $impersonatedAddress = "") {
    # Creates and returns an ExchangeService object to be used to access mailboxes
    $Script:ewsService = EWSAuth -Environment $Environment -Token $Script:Token -EwsUri $Script:EwsEndpoint
    return $Script:ewsService
}

function EWSPropertyType($MAPIPropertyType) {
    # Return the EWS property type for the given MAPI Property value

    switch ([Convert]::ToInt32($MAPIPropertyType, 16)) {
        0x0 { return $Null }
        0x1 { return $Null }
        0x2 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short }
        0x1002 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ShortArray }
        0x3 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer }
        0x1003 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::IntegerArray }
        0x4 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Float }
        0x1004 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::FloatArray }
        0x5 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Double }
        0x1005 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::DoubleArray }
        0x6 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Currency }
        0x1006 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CurrencyArray }
        0x7 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ApplicationTime }
        0x1007 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ApplicationTimeArray }
        0x0A { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Error }
        0x0B { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean }
        0x0D { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Object }
        0x100D { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ObjectArray }
        0x14 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long }
        0x1014 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::LongArray }
        0x1E { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String }
        0x101E { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray }
        0x1F { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String }
        0x101F { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::StringArray }
        0x40 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime }
        0x1040 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTimeArray }
        0x48 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CLSID }
        0x1048 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::CLSIDArray }
        0x102 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary }
        0x1102 { return [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::BinaryArray }
    }
    Write-Verbose "Couldn't match MAPI property type"
    return $Null
}

function InitPropList() {
    # We need to convert the properties to EWS extended properties
    if ($null -eq $script:itemPropsEws) {
        Write-Verbose "Building list of properties to retrieve"
        $script:property = @()
        foreach ($property in $ViewProperties) {
            $propDef = $null

            if ($property.StartsWith("{")) {
                # Property definition starts with a GUID, so we expect one of these:
                # {GUID}/name/mapiType - named property
                # {GUID]/id/mapiType   - MAPI property (shouldn't be used when accessing named properties)

                $propElements = $property -Split "/"
                if ($propElements.Length -eq 2) {
                    # We expect three elements, but if there are two it most likely means that the MAPI property Id includes the Mapi type
                    if ($propElements[1].Length -eq 8) {
                        $propElements += $propElements[1].Substring(4)
                        $propElements[1] = [Convert]::ToInt32($propElements[1].Substring(0, 4), 16)
                    }
                }
                $guid = New-Object Guid($propElements[0])
                $propType = EWSPropertyType($propElements[2])

                try {
                    $propDef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($guid, $propElements[1], $propType)
                } catch {
                    Write-Error "Unable to define property definitions."
                }
            } else {
                # Assume MAPI property
                if ($property.ToLower().StartsWith("0x")) {
                    $property = $deleteProperty.SubString(2)
                }
                $propId = [Convert]::ToInt32($deleteProperty.SubString(0, 4), 16)
                $propType = EWSPropertyType($deleteProperty.SubString(5))

                try {
                    $propDef = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($propId, $propType)
                } catch {
                    Write-Error "Unable to define property definitions."
                }
            }

            if ($null -ne $propDef) {
                $script:property += $propDef
                Write-Verbose "Added property $property to list of those to retrieve"
            } else {
                Write-Host "Failed to parse (or convert) property $property" -ForegroundColor Red
            }
        }
    }
}

$script:excludedProperties = @("Schema", "Service", "IsDirty", "IsAttachment", "IsNew")
$script:itemRetryCount = @{}

function InitLists() {
    $genericItemIdList = [System.Collections.Generic.List``1].MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId])
    $script:ItemsToDelete = [Activator]::CreateInstance($genericItemIdList)
}

function ProcessItem( $item ) {
    # We have found an item, so this function handles any processing
    $script:RequiredPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    #PR_POLICY_TAG 0x3019
    $PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
    #PR_RETENTION_DATE 0x301C
    $RetentionDate = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
    $script:RequiredPropSet.Add($PolicyTag)
    $script:RequiredPropSet.Add($RetentionDate)
    $item.Load($script:RequiredPropSet)

    $itemResult = New-Object PSObject -Property @{ InternetMessageId=$item.InternetMessageId; Sender=$item.Sender; ReceivedBy=$item.ReceivedBy; Id=$item.Id; ItemClass=$item.ItemClass; Subject=$item.Subject; DateTimeCreated=$item.DateTimeCreated; Folder=$folderPath; MailboxType=$Script:MailboxType; RetentionPeriod = $item.RetentionDate; PolicyTag = $item.PolicyTag }
    $itemResult | Export-Csv -Path $ResultsFile -NoTypeInformation -Append
}

function GetFolder() {
    # Return a reference to a folder specified by path

    $RootFolder, $FolderPath, $Create = $args[0]

    if ($null -eq  $RootFolder) {
        #LogVerbose "GetFolder called with null root folder"
        return $null
    }

    if ($FolderPath.ToLower().StartsWith("wellknownfoldername")) {
        # Well known folder, so bind to it directly
        $wkf = $FolderPath.SubString(20)
        #LogVerbose "Attempting to bind to well known folder: $wkf"
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf, $mbx )
        $Folder = ThrottledFolderBind($folderId)
        return $Folder
    }

    $Folder = $RootFolder
    if ($FolderPath -ne '\') {
        $PathElements = $FolderPath -split '\\'
        for ($i=0; $i -lt $PathElements.Count; $i++) {
            if ($PathElements[$i]) {
                $View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2, 0)
                $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
                $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
                $FolderResults = $Null
                try {
                    $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                } catch {
                    Write-Error "Unable to locate folders."
                }
                if ($null -eq $FolderResults) {
                    if (Throttled) {
                        try {
                            $FolderResults = $Folder.FindFolders($SearchFilter, $View)
                        } catch {
                            Write-Error "Unable to locate folders."
                        }
                    }
                }

                if ($null -eq $FolderResults) {
                    return $null
                }

                if ($FolderResults.TotalCount -gt 1) {
                    # We have more than one folder returned... We shouldn't ever get this, as it means we have duplicate folders
                    $Folder = $null
                    Write-Host "Duplicate folders ($($PathElements[$i])) found in path $FolderPath" -ForegroundColor Red
                    break
                } elseif ( $FolderResults.TotalCount -eq 0 ) {
                    if ($Create) {
                        # Folder not found, so attempt to create it
                        $subfolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($RootFolder.Service)
                        $subfolder.DisplayName = $PathElements[$i]
                        try {
                            $subfolder.Save($Folder.Id)
                            #LogVerbose "Created folder $($PathElements[$i])"
                        } catch {
                            # Failed to create the subfolder
                            $Folder = $null
                            Write-Host "Failed to create folder $($PathElements[$i]) in path $FolderPath" -ForegroundColor Red
                            break
                        }
                        $Folder = $subfolder
                    } else {
                        # Folder doesn't exist
                        $Folder = $null
                        Write-Host "Folder $($PathElements[$i]) doesn't exist in path $FolderPath" -ForegroundColor Red
                        break
                    }
                } else {
                    $Folder = ThrottledFolderBind $FolderResults.Folders[0].Id $null $RootFolder.Service
                }
            }
        }
    }

    $Folder
}

function SearchMailbox() {
    $Script:ewsService = CreateService($Mailbox)
    if ($null -eq $Script:ewsService) {
        return
    }
    # Set our root folder
    if ($Archive) {
        $Script:MailboxType = "Archive"
        if ($SearchDumpster) {
            $ProcessSubfolders = $True
            $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRecoverableItemsRoot
        } else {
            $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot
        }
    } else {
        $Script:MailboxType = "Primary"
        if ($SearchDumpster) {
            $ProcessSubfolders = $True
            $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot
        } else {
            $rootFolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot
        }
    }

    if (!($IncludeFolderList)) {
        # No folders specified to search, so the entire mailbox will be searched
        $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId( $rootFolderId )
        $ProcessSubfolders = $true
        SearchFolder $FolderId
    } else {
        # Searching specific folders
        $rootFolder = ThrottledFolderBind $rootFolderId
        foreach ($includedFolder in $IncludeFolderList) {
            $folder = $null
            $folder = GetFolder($rootFolder, $includedFolder, $false)

            if ($folder) {
                $folderPath = GetFolderPath($folder)
                SearchFolder $folder.Id
            }
        }
    }
}

function Throttled() {
    # Checks if we've been throttled.  If we have, we wait for the specified number of BackOffMilliSeconds before returning

    if ([String]::IsNullOrEmpty($script:Tracer.LastResponse)) {
        return $false # Throttling does return a response, if we don't have one, then throttling probably isn't the issue (though sometimes throttling just results in a timeout)
    }

    $lastResponse = $script:Tracer.LastResponse.Replace("<?xml version=`"1.0`" encoding=`"utf-8`"?>", "")
    $lastResponse = "<?xml version=`"1.0`" encoding=`"utf-8`"?>$lastResponse"
    $responseXml = [xml]$lastResponse

    if ($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value.Name -eq "BackOffMilliseconds") {
        # We are throttled, and the server has told us how long to back off for
        Write-Host "Throttling detected, server requested back off for $($responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text") milliseconds" Yellow
        Start-Sleep -Milliseconds $responseXml.Trace.Envelope.Body.Fault.detail.MessageXml.Value."#text"
        Write-Host "Throttling budget should now be reset, resuming operations" Gray
        return $true
    }
    return $false
}

function ThrottledFolderBind() {
    param (
        [Microsoft.Exchange.WebServices.Data.FolderId]$folderId,
        $propSet = $null,
        $exchangeService = $null
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    $folder = $null
    if ($null -eq $exchangeService) {
        $exchangeService = $Script:ewsService
    }

    try {
        if ($null -eq $propSet) {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
        } else {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propSet)
        }
        if (!($null -eq $folder)) {
            Write-Verbose "Successfully bound to folder $folderId"
        }
        return $folder
    } catch {
        Write-Error "Unable to bind to the $($folderId) folder."
    }

    if (Throttled) {
        try {
            if ($null -eq $propSet) {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId)
            } else {
                $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $folderId, $propSet)
            }
            if (!($null -eq $folder)) {
                Write-Verbose "Successfully bound to folder $folderId"
            }
            return $folder
        } catch {
            Write-Error "Unable to bind to the $($folderId) folder."
        }
    }

    # If we get to this point, we have been unable to bind to the folder
    Write-HostLog"FAILED to bind to folder $folderId"
    return $null
}

function GetFolderPath($Folder) {
    # Return the full path for the given folder

    # We cache our folder lookups for this script
    if (!$script:folderCache) {
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive
        # We use a .Net Dictionary object instead
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
    }

    $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId)
    $parentFolder = ThrottledFolderBind $Folder.Id $propSet $Folder.Service
    $folderPath = $Folder.DisplayName
    $parentFolderId = $Folder.Id
    while ($parentFolder.ParentFolderId -ne $parentFolderId) {
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId)) {
            try {
                $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId]
            } catch {
                Write-Error "Unable to find the parent folder."
            }
        } else {
            $parentFolder = ThrottledFolderBind $parentFolder.ParentFolderId $propSet $Folder.Service
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder)
        }
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath
        $parentFolderId = $parentFolder.Id
    }
    return $folderPath
}

function GetWellKnownFolderPath($WellKnownFolder) {
    if (!$script:wellKnownFolderCache) {
        $script:wellKnownFolderCache = @{}
    }

    if ($script:wellKnownFolderCache.ContainsKey($WellKnownFolder)) {
        return $script:wellKnownFolderCache[$WellKnownFolder]
    }

    $folder = $null
    $folderPath = $null
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:ewsService, $WellKnownFolder)
    if ($folder) {
        $folderPath = GetFolderPath($folder)
        #LogVerbose "GetWellKnownFolderPath: Path for $($WellKnownFolder): $folderPath"
    }
    $script:wellKnownFolderCache.Add($WellKnownFolder, $folderPath)
    return $folderPath
}

function IsFolderExcluded() {
    # Return $true if folder is in the excluded list

    param ($folderPath)

    # To support localization, we need to handle WellKnownFolderName enumeration
    # We do this by putting all our excluded folders into a hash table, and checking that we have the full path for any well known folders (which we retrieve from the mailbox)
    if ($null -eq $script:excludedFolders) {
        # Create and build our hash table
        $script:excludedFolders = @{}

        if ($ExcludeFolderList) {
            #LogVerbose "Building folder exclusion list"#: $($ExcludeFolderList -join ',')"
            foreach ($excludedFolder in $ExcludeFolderList) {
                $excludedFolder = $excludedFolder.ToLower()
                $wkfStart = $excludedFolder.IndexOf("wellknownfoldername")
                #LogVerbose "Excluded folder: $excludedFolder"
                if ($wkfStart -ge 0) {
                    # Replace the well known folder name with its full path
                    $wkfEnd = $excludedFolder.IndexOf("\", $wkfStart)-1
                    if ($wkfEnd -lt 0) { $wkfEnd = $excludedFolder.Length }
                    $wkf = $null
                    $wkf = $excludedFolder.SubString($wkfStart+20, $wkfEnd - $wkfStart - 19)

                    $wellKnownFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wkf
                    $wellKnownFolderPath = GetWellKnownFolderPath($wellKnownFolder)

                    $excludedFolder = $excludedFolder.Substring(0, $wkfStart) + $wellKnownFolderPath + $excludedFolder.Substring($wkfEnd+1)
                    #LogVerbose "Path of excluded folder: $excludedFolder"
                }
                $script:excludedFolders.Add($excludedFolder, $null)
            }
        }
    }

    return $script:excludedFolders.ContainsKey($folderPath.ToLower())
}

function SearchFolder( $FolderId ) {
    # Bind to the folder and show which one we are processing
    $folder = $null
    if($null -notlike $Script:Token) {
        CheckTokenExpiry -Environment $Environment -Token ([ref]$Script:Token) -EWSService ([ref]$Script:ewsService) -ApplicationInfo $Script:applicationInfo -EWSOnlineURL $Script:ewsOnlineURL -EWSOnlineScope $Script:ewsOnlineScope -AzureADEndpoint $azureADEndpoint
    }
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:ewsService, $FolderId)

    if ($null -eq $folder) { return }

    $folderPath = GetFolderPath($folder)

    if (IsFolderExcluded($folderPath)) {
        return
    }

    InitLists

    Write-Host "Searching the $($folderPath) for items." -ForegroundColor Gray

    # Search the folder for any matching items
    $pageSize = 100 # We will get details for up to 100 items at a time
    $moreItems = $true

    # Configure ItemView
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
    $view.Offset = 0
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow

    # Configure the search filter using the provided criteria
    $filters = @()

    if (![String]::IsNullOrEmpty($MessageClass)) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, $MessageClass)
    }

    if (![String]::IsNullOrEmpty($Subject)) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject)
    }

    if (![String]::IsNullOrEmpty($Sender)) {
        $senderEmailAddress = New-Object Microsoft.Exchange.WebServices.Data.EmailAddress($Sender)
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, $senderEmailAddress)
    }

    if (![String]::IsNullOrEmpty($MessageId)) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId, $MessageId)
    }

    if (![string]::IsNullOrEmpty($MessageBody)) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Body, $MessageBody)
    }

    # Add filter(s) for creation time
    if ( $CreatedAfter ) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $CreatedAfter)
    }
    if ( $CreatedBefore ) {
        $filters += New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $CreatedBefore)
    }

    # Create the search filter
    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
    foreach ($filter in $filters) {
        #LogVerbose([string]::Format("Adding search filter: {0}.", $filter.Value))
        $searchFilter.Add($filter)
    }

    # Perform the search and display the results
    while ($moreItems) {
        if($null -notlike $Script:Token) {
            CheckTokenExpiry -Environment $Environment -Token ([ref]$Script:Token) -EWSService ([ref]$Script:ewsService) -ApplicationInfo $Script:applicationInfo -EWSOnlineURL $Script:ewsOnlineURL -EWSOnlineScope $Script:ewsOnlineScope -AzureADEndpoint $azureADEndpoint
        }
        if($null -notlike $searchFilter) {
            $results = $Script:ewsService.FindItems( $FolderId, $searchFilter, $view )
        }
        else {
            $results = $Script:ewsService.FindItems( $FolderId, $view )
        }
        
        if ($results.Count -gt 0) {
            foreach ($item in $results.Items) {
                ProcessItem $item
            }
        }

        $moreItems = $results.MoreAvailable
        $view.Offset += $pageSize
    }

    # Now search subfolders
    if ($ProcessSubfolders) {
        $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
        $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
        foreach ($subFolder in $folder.FindFolders($view)) {
            SearchFolder $subFolder.Id $folderPath
        }
    }
}
}
process {}
end {
    $Date = [DateTime]::Now
    $Script:StartTime = '{0:MM/dd/yyyy HH:mm:ss}' -f $Date
    $ResultsFile = "$OutputPath\$Mailbox-SearchResults-$('{0:MMddyyyyHHmms}' -f $Date).csv"

    # Load EWS Managed API
    if (!(LoadEWSManagedAPI)) {
        Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
        exit
    }

    $script:searchResults = @()

    # Process as single mailbox
    $loggerParams = @{
        LogDirectory             = $OutputPath
        LogName                  = "EwsSearchMailbox-$((Get-Date).ToString("yyyyMMddhhmmss"))-Debug"
        AppendDateTimeToFileName = $false
        ErrorAction              = "SilentlyContinue"
    }

    $Script:Logger = Get-NewLoggerInstance @loggerParams

    SetWriteHostAction ${Function:Write-HostLog}
    SetWriteVerboseAction ${Function:Write-VerboseLog}
    SetWriteWarningAction ${Function:Write-HostLog}

    $cloudService = Get-CloudServiceEndpoint $AzureEnvironment

    # Define the endpoints that we need for the various calls to the Azure AD Graph API and EWS
    $Script:ewsOnlineScope = "$($cloudService.ExchangeOnlineEndpoint)/.default"
    $autoDSecureName = $cloudService.AutoDiscoverSecureName
    $azureADEndpoint = $cloudService.AzureADEndpoint

    $Script:applicationInfo = @{
        "TenantID" = $OAuthTenantId
        "ClientID" = $OAuthClientId
    }
    #Verify EWS URL
    if($Environment -eq "Onprem") {
        if([string]::IsNullOrEmpty($EwsUrl)) {
            Write-Warning "The Exchange Web Services URL must be provided for on-premises mailboxes."
            exit
        }
        else{
            if(-not($EwsUrl.EndsWith("EWS/Exchange.asmx"))) {
                if(-not($EwsUrl.EndsWith("/"))) {
                    $EwsUrl = "$($EwsUrl)/"
                }
            }
            $Script:EwsEndpoint = "$($EwsUrl)EWS/Exchange.asmx"
        }
    }
    else {
        $Script:EwsEndpoint = "$($cloudService.ExchangeOnlineEndpoint)/EWS/Exchange.asmx"
    }

    #Get OAuth token for authentication
    if(-not([string]::IsNullOrEmpty($OAuthClientId))){
        Get-OAuthToken -AppScope $Scope -Api EWS
    }
    else {
        if($Environment -eq "Online") {
            Write-Warning "OAuth parameters are required for Exchange Online mailboxes."
            exit
        }
        elseif($null -like $credential) {
            $credential = Get-Credential -Message "Credentials to access mailbox using EWS"
        }
    }

    # Perform the search
    SearchMailbox
}

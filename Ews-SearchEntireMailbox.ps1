<#
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
//
// SYNTAX:
// This example searches a cloud mailbox for the Subject "Script" within the inbox between the date 30 Sep thru 01 Nov using an impersonation account.
// .\Ews-SearchByFields.ps1 -Subject Script -MailboxName thanos@thejimmartin.com -FolderName Inbox -MailboxLocation Cloud -UseImpersonation:$true -OlderThan "11/1/2022" -LaterThan "9/30/2022"
//
// This example searches an on-premises mailbox for message from venom@thejimmartin.com sent before 01 Nov in the Inbox using an impersonation account.
// .\Ews-SearchByFields.ps1 -MailboxName groot@thejimmartin.com -FolderName Inbox -MailboxLocation OnPremises -UseImpersonation:$true -EwsURL owa.thejimmartin.com -OlderThan "11/1/2022" -Sender venom@thejimmartin.com
#>

param(
    [Parameter(Mandatory = $false)] [string] $Subject,
    [Parameter(Mandatory = $false)] [System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory = $true, HelpMessage="Mailbox to search")] [string] $MailboxName,
    [Parameter(Mandatory = $false, HelpMessage="Sender to search for in items")] [string] $Sender,
    [Parameter(Mandatory = $false, HelpMessage="The start date for your search criteria. All messages older than this date will be deleted.")] [datetime] $OlderThan,
    [Parameter(Mandatory = $false, HelpMessage="The start date for your search criteria. All messages newer than this date will be deleted.")] [datetime] $LaterThan,
    [Parameter(Mandatory = $false, HelpMessage="Account used has impersonation rights")] [boolean] $UseImpersonation=$false,
    [Parameter(Mandatory = $false, HelpMessage="Enables EWS trace logging")] [boolean] $EnableLogging=$false,
    [Parameter(Mandatory = $false, HelpMessage="Location of the mailbox")] [ValidateSet("OnPremises", "Cloud")] [string]$MailboxLocation="Cloud",
    [Parameter(Mandatory = $false, HelpMessage="Use OAuth for authentication")] [boolean] $OAuth= $(if($MailboxLocation -eq "Cloud") {$true} else {$false}),
    [Parameter(Mandatory = $false, HelpMessage="EWS namespace for on-premises environment (ex: ews.contoso.com)")] [string] $EwsURL = $(if($MailboxLocation -eq "Cloud"){"outlook.office365.com"} else {throw "-EwsURL must be passed for on-premises mailbox."}),
    [Parameter(Mandatory = $false, HelpMessage="Application permission type of either Delegated or Application")] [ValidateSet("Delegated", "Application")] [String]$ApplicationPermission="Delegated",
    [Parameter(Mandatory = $false, HelpMessage="Mailbox being accessed is an archive mailbox")] [boolean] $Archive=$false
)

function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
}  

function Enable-TraceHandler(){
    $sourceCode = @"
        public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
        {
            public System.String LogFile {get;set;}
            public void Trace(System.String traceType, System.String traceMessage)
            {
                System.IO.File.AppendAllText(this.LogFile, traceMessage);
            }
        }
"@    
    Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $ewsDLL #$Script:EWSDLL
    $TraceListener = New-Object ewsTraceListener
    return $TraceListener
}

function Get-OAuthToken{
    #Change the AppId, AppSecret, and TenantId to match your registered application
    $AppId = "6a93c8c4-9cf6-4efe-a8ab-9eb178b8dff4"
    $AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    $TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
    #Build the URI for the token request
    $Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        client_id     = $AppId
        scope         = "https://$EwsURL/.default"
        client_secret = $AppSecret
        grant_type    = "client_credentials"
    }
    $TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
    #Unpack the access token
    $Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
    return $Token
}

function Get-DelegatedOAuthToken {
    #Check and install Microsoft Authentication Library module
    if(!(Get-Module -Name MSAL.PS -ListAvailable -ErrorAction Ignore)){
        try { 
            #Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module -Name MSAL.PS -Repository PSGallery -Force
        }
        catch {
            Write-Warning "Failed to install the Microsoft Authentication Library module."
            exit
        }
        try {
            Import-Module -Name MSAL.PS
        }
        catch {
            Write-Warning "Failed to import the Microsoft Authentication Library module."
        }
    }
    $ClientID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    $RedirectUri = "ms-appx-web://Microsoft.AAD.BrokerPlugin/d3590ed6-52b3-4102-aeff-aad2292ab01c"
    $Token = Get-MsalToken -ClientId $ClientID -RedirectUri $RedirectUri -Scopes "https://$EwsURL/EWS.AccessAsUser.All" -Interactive
    #$OAuthToken = "Bearer {0}" -f $Token.AccessToken
    return $Token.AccessToken
}

function Find-MailboxItem {
    param(
    [Parameter(Mandatory = $false)] $Folder,
    [Parameter(Mandatory = $false)] $Root
    )
    $ItemSearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        #Create a search filter for a blank subject
        #$ItemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject)
        #$ItemNotSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($ItemSearchFilter)
        #$ItemSearchFilterCollection.Add($ItemNotSearchFilter)
        if($Subject -notlike $null) {
            #Create a search filter for a subject
            $ItemSubjectFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $Subject)
            $ItemSearchFilterCollection.Add($ItemSubjectFilter)
        }
        if($Sender -notlike $null) {
            #Create a search filter for a sender
            $ItemSenderFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender,$Sender)
            $ItemSearchFilterCollection.Add($ItemSenderFilter)
        }
        if($OlderThan -notlike $null) {
            $TempStartDate = [datetime]$OlderThan
            $TempStartDate = $TempStartDate.ToUniversalTime()
            $SearchStartDate = '{0:yyyy-MM-ddThh:mm:ssZ}' -f $TempStartDate
            #Create a search filter for recieved time
            $ItemReceiveTimeLessFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, $SearchStartDate)
            $ItemSearchFilterCollection.Add($ItemReceiveTimeLessFilter)
        }
        if($LaterThan -notlike $null) {
            $TempStartDate = [datetime]$LaterThan
            $TempStartDate = $TempStartDate.ToUniversalTime()
            $SearchStartDate = '{0:yyyy-MM-ddThh:mm:ssZ}' -f $TempStartDate
            #Create a search filter for recieved time
            $ItemReceiveTimeGreaterFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, $SearchStartDate)
            $ItemSearchFilterCollection.Add($ItemReceiveTimeGreaterFilter)
        }
        $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)     
        $fiItems = $null   
        #$OutputPath = Get-Location   
        do{   
            $psPropset= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            #Perform the search using the combined search filters
            $fiItems = $Folder.findItems($ItemSearchFilterCollection,$ivItemView)
            if($fiItems.Items.Count -gt 0){  
                [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)    
                foreach($Item in $fiItems.Items){   
                    $MailboxObj = New-Object PSObject -Property @{ Mailbox=$MailboxName; Folder=$ffFolder.DisplayName; Subject=$Item.Subject; From=$Item.From; DateTimeReceived=$Item.DateTimeReceived; Size=$Item.Size; DateTimeSent=$Item.DateTimeSent; HasAttachments=$Item.HasAttachments; MessageClass=$Item.ItemClass; RootFolder=$Root};
                    #Write-Output $MailboxObj
            
                    $MailboxObj | Export-Csv $OutputFile -Append -NoTypeInformation
                }  
            }  
            $ivItemView.Offset += $fiItems.Items.Count
        }
        while($fiItems.MoreAvailable -eq $true)
        #endregion

}

#region LoadEwsManagedAPI
#Check for EWS Managed API, exit if missing
$ewsDLL = (($(Get-ItemProperty -ErrorAction Ignore -Path Registry::$(Get-ChildItem -ErrorAction Ignore -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' |Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
if($ewsDLL -notlike $null) {
    $ewsDLL = $ewsDLL+"Microsoft.Exchange.WebServices.dll"
}
else {
    $ScriptPath = Get-Location
    $ewsDLL = "$ScriptPath\Microsoft.Exchange.WebServices.dll"
    Unblock-File $ewsDLL -Confirm:$false
}
if (Test-Path $ewsDLL) {
    Import-Module $ewsDLL
}
else {
    Write-Warning "This script requires the EWS Managed API 1.2 or later."
    exit
}
#endregion

$WellKnownFolderNames = @("ArchiveDeletedItems",
"ArchiveMsgFolderRoot",
"ArchiveRecoverableItemsDeletions",
"ArchiveRecoverableItemsPurges",
"ArchiveRecoverableItemsRoot",
"ArchiveRecoverableItemsVersions",
"ArchiveRoot",
"Calendar",
"Conflicts",
"Contacts",
"ConversationHistory",
"DeletedItems",
"Drafts",
"Inbox",
"Journal",
"JunkEmail",
"LocalFailures",
"MsgFolderRoot",
"Notes",
"Outbox",
"PublicFoldersRoot",
"QuickContacts",
"RecipientCache",
"RecoverableItemsDeletions",
"RecoverableItemsPurges",
"RecoverableItemsRoot",
"RecoverableItemsVersions",
"Root",
"SearchFolders",
"SentItems",
"ServerFailures",
"SyncIssues",
"Tasks",
"ToDoSearch",
"VoiceMail"
)

$OutputPath = Get-Location
$OutputFile = "$OutputPath\$MailboxName-SearchResults.csv"
if(Get-Item $OutputFile -ErrorAction Ignore) {
    Remove-Item $OutputFile -Confirm:$False -ErrorAction Ignore
}

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion) 
$service.HttpHeaders.Clear()
#region Get credentials
if($OAuth) {
    if($ApplicationPermission -eq "Application") {
        $Token = Get-OAuthToken
    }
    else {
        $Token = Get-DelegatedOAuthToken
    }
    $OAuthToken = "Bearer {0}" -f $Token
    $service.HttpHeaders.Add("Authorization", " $($OAuthToken)")
}
else {
    $psCred = Get-Credential  
    $creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
    $service.Credentials = $creds
}
#endregion

if($MailboxLocation -eq "OnPremises" -and $EwsURL -like $null) {
    $service.AutodiscoverUrl($MailboxName,{$true})
}
else {
    $service.Url = "https://$EwsURL/ews/exchange.asmx"    
}
$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
if($UseImpersonation -or $ApplicationPermission -eq "Application") {
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
}
$service.UserAgent = "EwsPowerShellScript"
if($EnableLogging) {
    Write-Host "EWS trace logging enabled" -ForegroundColor Cyan
    $service.TraceEnabled = $True
    $TraceHandlerObj = Enable-TraceHandler
    $OutputPath = Get-Location
    $TraceHandlerObj.LogFile = "$OutputPath\$MailboxName-TraceLog.log"
    $service.TraceListener = $TraceHandlerObj
}

if($Archive) {
$SearchFolderNames = @("ArchiveDeletedItems",
"ArchiveMsgFolderRoot",
"ArchiveRecoverableItemsDeletions",
"ArchiveRecoverableItemsPurges",
"ArchiveRecoverableItemsRoot",
"ArchiveRecoverableItemsVersions",
"MsgFolderRoot",
"RecoverableItemsDeletions",
"RecoverableItemsPurges",
"RecoverableItemsRoot",
"RecoverableItemsVersions"
)
}
else {
    $SearchFolderNames = @("MsgFolderRoot",
"RecoverableItemsDeletions",
"RecoverableItemsPurges",
"RecoverableItemsRoot",
"RecoverableItemsVersions"
)
}
foreach($FolderName in $SearchFolderNames) {
    #Define Extended properties
    $PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
    $folderidcnt = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderName,$MailboxName)  
    #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
    $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
    #Deep Transval will ensure all folders in the search path are returned  
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
    $psPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $PR_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
    $PR_DELETED_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
    $PR_Folder_Path = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
    #Add Properties to the  Property Set  
    $psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED);  
    $psPropertySet.Add($PR_Folder_Path);  
    $fvFolderView.PropertySet = $psPropertySet;  
    #The Search filter will exclude any Search Folders  
    $sfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
    $fiResult = $null  
    #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
    do {  
        $fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
        foreach($ffFolder in $fiResult.Folders){
            $SearchFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$ffFolder.Id)
            Write-Host "Searching the $($ffFolder.DisplayName) folder of $MailboxName for the subject `'$Subject`'..." -ForegroundColor Cyan
            Find-MailboxItem -Folder $SearchFolder -Root $FolderName
        
    } 
    $fvFolderView.Offset += $fiResult.Folders.Count
}while($fiResult.MoreAvailable -eq $true)  
}
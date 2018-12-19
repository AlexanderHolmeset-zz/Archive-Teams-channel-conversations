
### Team conversation archiver                  ###

### Version 1.0                                 ###

### Author: Alexander Holmeset                  ###

### Twitter: twitter.com/alexholmeset           ###

### Blog: alexholmeset.blog                     ###

#Description:
#You specify a Group/team object ID, then the script archives all conversations in every channel for this team in a HTML file.
#Have in mind information protection policies like GDPR when working on information like this.


#Parameters
#
$resourceURI = "https://graph.microsoft.com"
$authority = "https://login.microsoftonline.com/common"
#Enter your own clientID and redirect URI.Take a look at Ståle Hansen blogpost to see how to create your own clientid and redirect URI.
#https://msunified.net/2018/12/12/post-at-microsoftteams-channel-chat-message-from-powershell-using-graph-api/
#
$clientId = "3587e50e-98f4-4bbe-86ff-f94c2056a7e8"
$redirectUri = "https://login.microsoftonline.com/M365x792147.onmicrosoft.com/oauth2"


#Remove commenting on username and password if you want to run this without a prompt.
#$Office365Username='user@domain'
#$Office365Password='VeryStrongPassword' 


#pre requisites
try {
$AadModule = Import-Module -Name AzureAD -ErrorAction Stop -PassThru
}
catch {
throw 'Prerequisites not installed (AzureAD PowerShell module not installed)'
}
$adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
$adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
 
##option without user interaction
if (([string]::IsNullOrEmpty($Office365Username) -eq $false) -and ([string]::IsNullOrEmpty($Office365Password) -eq $false))
{
$SecurePassword = ConvertTo-SecureString -AsPlainText $Office365Password -Force
#Build Azure AD credentials object
$AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList $Office365Username,$SecurePassword
# Get token without login prompts.
$authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceURI, $clientid, $AADCredential);
}
else
{
# Get token by prompting login window.
$platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
$authResult = $authContext.AcquireTokenAsync($resourceURI, $ClientID, $RedirectUri, $platformParameters)
}
$accessToken = $authResult.result.AccessToken

 
#Group/team object ID. 
$TeamID = 'f8e68301-8d91-4360-b033-eeb78ca2077f'

#Where to store the HTML file:
$Storage = 'c:\temp\test.html'

#Gets all channels in a Team
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$TeamChannels = $myprofile.value | Select-Object ID,DisplayName

"
" | Out-File $Storage
  
foreach($Channel in $TeamChannels) {

#Gets all root messages/conversations in a channel.
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels/"+$channel.id+"/messages"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$ChannelMessages = $myprofile.value | Select-Object Body,From,ID,attachments,createdDateTime | Sort-Object
$Channeldisplayname = $channel.displayName


"<br>
 ------------------------------<br>
<br>
$channeldisplayname<br>
<br>" | Out-File -Append $Storage

foreach($channelmessage in $ChannelMessages){


$channelmessagedisplayname = (($channelmessage.from).user).displayname
$channelmessagecontent = ($channelmessage.body).content
$channelmessageattachment = ($channelmessage.attachments).contenturl
$ChannelmessagecreatedDateTime = $channelmessage.createdDateTime

"<br>
      ***************************<br>
      $channelmessagecreatedDateTime<br>
      $channelmessagedisplayname<br>
      $channelmessagecontent<br>
      $channelmessageattachment<br>
<br>" | out-file -Append $Storage

#Gets all replies to a root message in a channel.
$apiUrl = "https://graph.microsoft.com/beta/teams/$TeamID/channels/"+$channel.id+"/messages/"+$ChannelMessage.id+"/replies"
$myProfile = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get
$replies = $myprofile.value | Select-Object body,from,attachments,createdDateTime | Sort-Object

foreach($reply in $replies){

$replydisplayname = (($reply.from).user).displayname
$replycontent = ($reply.body).content
$replyattachment = ($reply.attachments).contenturl
$replycreatedDateTime = $reply.createdDateTime


"<br>
            $replycreatedDateTime<br>
            $replydisplayname<br>
            $replycontent<br>
            $replyattachment<br>
<br>" | out-file  -Append $Storage



}



}


}

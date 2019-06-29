#Connect to Exchange Online using your Office 365 administrative credentials
$credential = Get-Credential
Install-Module MsOnline
Connect-MsolService -Credential $credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

#Connect to SharePoint Online
#Connect SPOService -Url https://<SP Admin Center>.sharepoint.com -credential $cred

#Connect to AzureAD
Connect-MsolService -Credential $cred

# Set the user variable with the user that is to be offboarded
$Username = Read-Host -Prompt "Email of user leaving"


#Initializing Variables
$User = Get-AzureADUser -ObjectId $Username
$Mailbox = Get-Mailbox | Where {$_.PrimarySmtpAddress -eq $username}
$Manager = Get-AzureADUserManager -ObjectId $user.ObjectId
$OutOfOfficeBody = @"
Hello
Please Note I am no longer work for Blue Spruce Capital anymore.
Please contact $($Manager.DisplayName) $($Manager.UserPrincipalName) for any questions.
Thanks!
"@

#Set Sign in Blocked
Set-AzureADUser -ObjectId $user.ObjectId -AccountEnabled $false

#Disconnect Existing Sessions
Revoke-SPOUserSession -User $Username -confirm:$False
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId

#Forward e-mails to manager
Set-Mailbox $Mailbox.Alias -ForwardingAddress $Manager.UserPrincipalName -DeliverToMailboxAndForward $False -HiddenFromAddressListsEnabled $true

#Set Out Of Office
Set-MailboxAutoReplyConfiguration -Identity $Mailbox.Alias -ExternalMessage $OutOfOfficeBody -InternalMessage $OutOfOfficeBody -AutoReplyState Enabled

#Cancel meetings organized by this user
Remove-CalendarEvents -Identity $Mailbox.Alias -CancelOrganizedMeetings -confirm:$False

#RemoveFromDistributionGroups
$DistributionGroups= Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains "$Username"}

foreach( $dg in $DistributionGroups)
	{
	Remove-DistributionGroupMember $dg.name -Member $Username -Confirm:$false
	}

#Re-Assign Office 365 Group Ownership
$Office365GroupsOwner = Get-UnifiedGroup | where { (Get-UnifiedGroupLinks $_.Alias -LinkType Owners| foreach {$_.name}) -contains $mailbox.Alias}
$NewManagerGroups = @()
foreach($GRP in $Office365GroupsOwner)
	{
	$Owners = Get-UnifiedGroupLinks $GRP.Alias -LinkType Owners
	if ($Owners.Count -le 1)
		{
		#Our user is the only owner
		Add-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Members -Links $Manager.UserPrincipalName
		Add-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Owners -Links $Manager.UserPrincipalName
		$NewManagerGroups += $GRP
		Remove-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Owners -Links $Username -Confirm:$false
		Remove-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Members -Links $Username -Confirm:$false
		}
	else
		{
		#There Are Other Owners
		Remove-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Owners -Links $Username -Confirm:$false
		}
	}

#Remove from Office 365 Groups
$Office365GroupsMember = Get-UnifiedGroup | where { (Get-UnifiedGroupLinks $_.Alias -LinkType Members | foreach {$_.name}) -contains $mailbox.Alias}
$NewMemberGroups = @()
foreach($GRP in $Office365GroupsMember)
	{
	$Members = Get-UnifiedGroupLinks $GRP.Alias -LinkType Members
	if ($Members.Count -le 1)
		{
		#Our user is the only Member
		Add-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Members -Links $Manager.UserPrincipalName
		$NewMemberGroups += $GRP
		Remove-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Members -Links $Username -Confirm:$false
		}
	else
		{
		#There Are Other Members
		Remove-UnifiedGroupLinks -Identity $GRP.Alias -LinkType Members -Links $Username -Confirm:$false
		}
	}

#Send OneDrive for Business Information to Manager
$OneDriveUrl = Get-PnPUserProfileProperty -Account $username | select PersonalUrl
Set-SPOUser $Manager.UserPrincipalName -Site $OneDriveUrl.PersonalUrl -IsSiteCollectionAdmin:$true

#Convert Mailbox to Shared
Set-Mailbox $Username -Type Shared

#Send Final E-mail to Manager

#BuildHTMLObjects

If ($DistributionGroups)
{
	$DGHTML = " The user has been removed from the following distribution lists

    " foreach( $dg in $DistributionGroups) { $DGHTML += "
    $($dg.PrimarySmtpAddress)
    " } $DGHTML += "

 "
}

If ($Office365GroupsOwner)
{
	$O365OwnerHTML = " The user was an owner, and was removed from the following groups

    " foreach($GRP in $Office365GroupsOwner) { $O365OwnerHTML += "
    $($GRP.PrimarySmtpAddress)
    " } $O365OwnerHTML += "

 "
}

If ($Office365GroupsMember)
{
	$O365MemberHTML = " The user was a member, and was removed from the following groups

    " foreach($GRP in $Office365GroupsMember) { $O365MemberHTML += "
    $($GRP.PrimarySmtpAddress)
    " } $O365MemberHTML += "

 "
}

If ($NewManagerGroups)
{
	$NewOwnerAlertHTML = " *Attention Required*   The user was the only owner of the following groups. Please verify if there is any content in those groups that is still needed, otherwise, archive the groups as per normal procedure

    " foreach($GRP in $NewManagerGroups) { $NewOwnerAlertHTML += "
    $($GRP.PrimarySmtpAddress)
    " } $NewOwnerAlertHTML += "

 "
}

If ($NewMemberGroups)
{
	$NewMemberAlertHTML = " *Attention Required*   The user was the only member of the following groups. Please verify if there is any content in those groups that is still needed, otherwise, contact the owner of the groups to be removed, or to archive the group

    " foreach($GRP in $NewMemberGroups) { $NewMemberAlertHTML += "
    $($GRP.PrimarySmtpAddress)
    " } $NewMemberAlertHTML += "

 "
}


$Subject = "User Offboarding Complete: $($User.UserPrincipalName)"
$ManagerEmailBody = @"
Hello $($Manager.DisplayName)

This is an automated e-mail from IT to let you know that the account $($User.UserPrincipalName)  has been de-activated as per normal standard procedure. All e-mails have been forwarded to you! $DGHTML $O365OwnerHTML $O365MemberHTML $NewOwnerAlertHTML $NewMemberAlertHTML

You have also been assigned ownership of the OneDrive for Business of the account. Please navigate to the following URL : $($OneDriveUrl.PersonalUrl) and save any important data within 30 days.

If you have any questions, please contact the IT Department.  Thank you!

"@

Send-MailMessage -To $Manager.UserPrincipalName -from someone@somecompany.com -Subject $Subject -Body ( $ManagerEmailBody | out-string ) -BodyAsHtml -smtpserver smtp.office365.com -usessl -Credential $cred -Port 587

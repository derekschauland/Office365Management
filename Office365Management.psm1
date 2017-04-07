<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.136
	 Created on:   	3/7/2017 10:45 AM
	 Created by:   	dschauland
	 Organization: 	
	 Filename:     	Office365Management.psm1
	-------------------------------------------------------------------------
	 Module Name: Office365Management
	===========================================================================
#>

<#
.EXTERNALHELP .\forward-o365mail.psm1-help.xml
#>

function connect-o365
{
	
	foreach ($session in (Get-PSSession | where { $_.configurationname -eq "Microsoft.exchange" }))
	{
			Write-Host "$session will be removed."
			$session | remove-pssession
		}
		
		$credential = Get-Credential -Message "Enter the O365 login for the customer you need to work with"
	
	$exsession = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri "https://ps.outlook.com/powershell" -Credential $credential -Authentication basic -AllowRedirection
	
	Export-PSSession -Session $exsession -OutputModule "ExchOnline" -AllowClobber -Force
	
	if (!(Get-Module ExchOnline))
	{
		Import-Module exchonline -DisableNameChecking | Write-Output $null
	}
	else
	{
		Get-Module exchonline | Remove-Module 	
	}
	
	Import-Module exchonline -DisableNameChecking | Write-Output $null
	
	
}

function enable-o365mailforward
{
	param (
		[string]$emailaddress,
		[string]$forwarddestination,
		[switch]$keepcopy
	)
	
	if(get-pssession | where { $_.configurationname -eq "Microsoft.Exchange" }])
	{
		$current = get-mailbox -identity $emailaddress | fl forwardingsmtpaddress, delivertomailboxandforward, identity
	}
	else
	{
		connect-o365
		
		$current = get-mailbox -identity $emailaddress | fl forwardingsmtpaddress, delivertomailboxandforward, identity
	}
	
	
	$yourcreds = Get-Credential -Message "Enter Your username and password for O365 to send a test mail to the changed mailbox"
	
	if ($current.forwardingsmtpaddress -eq $null)
	{
		Write-Host "Mail for $emailaddress is not being forwarded at this time"
		Write-Host "Will set mail to forward to $forwarddestination"
		
		if ($forwarddestination)
		{
			if ($keepcopy)
			{
				set-mailbox -identity $emailaddress -delivertomailboxandforward $true -forwardingsmtpaddress $forwarddestination
				Write-Host "Mail has been set to forward to $forwarddestination and a copy will be kept in the user's mailbox. Mailbox type set to Shared to remove usage of Office 365 License."
				
				Set-Mailbox $emailaddress -Type Shared
				
				Send-MailMessage -To $emailaddress -Subject "Test Email Please Reply" -SmtpServer smtp.office365.com -usessl -Port 587 -From $yourcreds.GetNetworkCredential().UserName -cc $yourcreds.GetNetworkCredential().username -Body "Please reply if received.`n`nThanks" -Credential $yourcreds
				
				Write-Host "Email testing the changes has been sent to $emailaddress and copied to $($yourcreds.getnetworkcredential().username)"
			}
			else
			{
				set-mailbox -identity $emailaddress -delivertomailboxandforward $false -forwardingsmtpaddress $forwarddestination
				Write-Host "Mail has been set to forward to $forwarddestination and a copy will not be kept in the user's mailbox. Mailbox type set to Shared to remove usage of Office 365 License."
				Set-Mailbox $emailaddress -Type Shared
				Send-MailMessage -To $emailaddress -Subject "Test Email Please Reply" -SmtpServer smtp.office365.com -usessl -Port 587 -From $yourcreds.GetNetworkCredential().UserName -cc $yourcreds.GetNetworkCredential().username -Body "Please reply if received.`n`nThanks" -Credential $yourcreds
				Write-Host "Email testing the changes has been sent to $emailaddress and copied to $($yourcreds.getnetworkcredential().username)"
			}
			
		}
		else
		{
			Write-Host "Please enter a forwarding address."	
		}
		
	}
	else
	{
		Write-Host "$emailaddress is having mail forwarded to $($current.fowardingsmtpaddress) currently - no changes made."
		
	}
	
	
}

function disable-o365mailforward
{
	param ([string]$emailaddress)
	
	$yourcreds = Get-Credential -Message "Enter Your username and password for O365 to send a test mail to the changed mailbox"
	
	if (Get-Mailbox -Identity $emailaddress)
	{
		$checkfwd = get-mailbox -identity $emailaddress | select forwardingsmtpaddress, delivertomailboxandforward
	}
	else
	{
		connect-o365
		
		$checkfwd = get-mailbox -identity $emailaddress | select forwardingsmtpaddress, delivertomailboxandforward
	}
	
	if ($checkfwd.forwardingsmtpaddress -ne  $null)
	{
		set-mailbox -identity $emailaddress -forwardingsmtpaddress $null -DeliverToMailboxAndForward $false
		Set-Mailbox $emailaddress -Type Regular 
		Write-Host "Mail forwarding for $emailaddress has been disabled - mailbox set back to Regular User Mailbox and will consume License"
		Send-MailMessage -To $emailaddress -Subject "Test Email Please Reply" -SmtpServer smtp.office365.com -usessl -Port 587 -From $yourcreds.GetNetworkCredential().UserName -cc $yourcreds.GetNetworkCredential().username -Body "Please reply if received.`n`nThanks" -Credential $yourcreds
		Write-Host "Email testing the changes has been sent to $emailaddress and copied to $($yourcreds.getnetworkcredential().username)"
	}
	else
	{
		Write-Host "No Forwarding to disable for $emailaddress"	
	}
	
}

function get-o365mailforward
{
	param (
		$emailaddress
	)
	
	if (get-mailbox -identity $emailaddress | Select-Object forwardingsmtpaddress, delivertomailboxandforward, identity)
	{
		$isforwarding = get-mailbox -identity $emailaddress | Select-Object forwardingsmtpaddress, delivertomailboxandforward, identity
	}
	else
	{
		connect-o365
		$isforwarding = get-mailbox -identity $emailaddress | Select-Object forwardingsmtpaddress, delivertomailboxandforward, identity
	}
	
	
	if ($isforwarding.forwardingsmtpaddress -ne $null)
	{
		Write-Host "Email for $emailaddress is being forwarded:`n $($isforwarding.forwardingsmtpaddress) is the current destination`n Is mailbox also keeping a copy: $($isforwarding.dilevertomailboxandforward)"	
	}
	else
	{
		Write-Host "Email for $emailaddress is not being forwarded to any other addresses."		
	}
}

function disable-o365access
{
	param ([string]$emailaddress,
	[string[]]$feature
	)
	
	switch ($feature) {
		"OWA" {
			Write-Host "Will disable OWA for $emailaddress"
			Set-CASMailbox $emailaddress -owaEnabled $false
		}
		"ActiveSync" {
			Write-Host "Will disable ActiveSync for $emailaddress"
			Set-CASMailbox $emailaddress -activesyncEnabled $false
		}
		"IMAP" {
			Write-Host "Will disable IMAP for $emailaddress"
			Set-CASMailbox $emailaddress -ImapEnabled $false
		}
		"POP" {
			Write-Host "Will disable POP for $emailaddress"
			Set-CASMailbox $emailaddress -popEnabled $false
		}
		"MAPI" {
			Write-Host "Will disable MAPI for $emailaddress"
			Set-CASMailbox $emailaddress -MAPIEnabled $false
		}
		default {
			Get-CASMailbox $emailaddress
		}
	}
	
}

function enable-o365access
{
	param ([string]$emailaddress,
		[string[]]$feature
	)
	
	switch ($feature)
	{
		"OWA" {
			Write-Host "Will enable OWA for $emailaddress"
			Set-CASMailbox $emailaddress -owaEnabled $true
		}
		"ActiveSync" {
			Write-Host "Will enable ActiveSync for $emailaddress"
			Set-CASMailbox $emailaddress -activesyncEnabled $true
		}
		"IMAP" {
			Write-Host "Will enable IMAP for $emailaddress"
			Set-CASMailbox $emailaddress -ImapEnabled $true
		}
		"POP" {
			Write-Host "Will enable POP for $emailaddress"
			Set-CASMailbox $emailaddress -popEnabled $true
		}
		"MAPI" {
			Write-Host "Will enable MAPI for $emailaddress"
			Set-CASMailbox $emailaddress -MAPIEnabled $true
		}
		default
		{
			Get-CASMailbox $emailaddress
		}
	}
	
}

function add-calendarreviewer
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$user 
	)
	BEGIN 
	{
		$mailbox = Get-Mailbox
		
	}
	PROCESS
	{
		$mailbox | ForEach-Object
		{
			if ((Get-MailboxFolderPermission $_":\Calendar" -User $user))
			{
				Write-Host "$user has access to $mailbox already"
			}
			else
			{
				Get-Mailbox |
				ForEach-Object {
					Add-MailboxFolderPermission $_":\Calendar" -User $user -AccessRights Reviewer -ErrorAction silentlycontinue
				}
			}
		}
		
	}
	END
	{
		
	}

		
}

function check-o365status
{
	param($username, $password,[switch]$full)
	
	$secpass = ConvertTo-SecureString $password -AsPlainText -Force
	
	$cred = New-Object System.Management.Automation.PSCredential($username, $secpass)
	
	$mysession = new-scsession -credential $cred
	
	if ($full)
	{
		Get-SCevent -scsession $mysession
	}
	else
	{
		get-scevent -scsession $mysession | Select-Object name, status	
	}
	
}

function update-o365primarymail
{
	param ([string[]]$emailaddress)
	
	$mailbox = get-mailbox $emailaddress
	
	foreach ($mbx in $mailbox)
	{
		if ($mbx.windowsemailaddress -match "onmicrosoft.com")
		{
			$oldemail = $mbx.windowsemailaddress
			
			$mbx | set-mailbox -windowsmailaddress $mbx.userprinicpalname
			
			Write-Host "The Office 365 account for $($mbx.displayname) has been modified to change $oldemail to $($mbx.windowsemailaddress)"
		}
		else
		{
			Write-Host "The Office 365 account for $($mbx.displayname) already has the correct Email address - no changes needed/made."	
		}
	}
}

Export-ModuleMember check-o365status, enable-o365mailforward, connect-o365, disable-o365mailforward, get-o365mailforward, disable-o365access, enable-o365access, add-calendarreviewer, update-o365primarymail




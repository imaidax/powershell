#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

$loggedinas = ''

function login
{
	#These two fields grab the text in the text boxes
	
	$username = $adminEmail.Text
	$password = $adminPassword.Text
	
	#This takes the password and converts it to a secure string (Despite it already being so) and making it usable for automatic login.
	$pass = ConvertTo-SecureString -String $password -AsPlainText -Force
	$cred = New-Object -TypeName System.Management.Automation.PSCredential($username, $pass)
	try
	{
		#Connect to Office365 Tenant
		$output.AppendText("`nAttemping login as $username to Office365 Tenant..")
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
		Import-PSSession -session $Session
		
		#Connect to Azure Tenant
		$output.AppendText("`nSuccess. Attempting signing into Azure Active Directory..")
		Connect-MsolService -Credential $cred
		
		#Succesful Login / Welcome & Titlebar Change 
		$you = get-user -identity $username | Select-Object DisplayName
		$you = $you -replace ".*=" -replace "}"
		$output.AppendText("`nLogin Succeeded.`nWelcome back $you")
		$formVERMicrosoftOffice36.Text = "VER Microsoft Office365 Tool. Logged in as $username"
		
		$loggedinas = $username
		# Clear user name if you need to login as someone else for whatever reason.
		$adminEmail.Text = ""
	}
	
	catch
	{
		$output.AppendText("`nLogin Failed. Please check your credentials.")
	}
	
	#Clears password after login or failure for easier reattempt
	
	$adminPassword.Text = ""
}

function check_calendar
{
	try
	{
		$folder = '{0}:\calendar' -f $CalendarEmail.Text
		
		$calendars = Get-MailboxFolderPermission -identity $folder | Select-Object User, AccessRights | Out-String
		$calendars = $calendars -replace "[{}]"
		
		$output.Text = $calendars
	}
	
	catch
	{
		$output.AppendText("Calendar not found. Check spelling or if it exists.")
	}
	
}

function set_calendar_permissions
{
	$calendar = '{0}:\calendar' -f $CalendarEmail.Text
	$user = '{0}' -f $UserEmail.Text
	if ($calendarSetRead.Checked -eq $true)
	{
		$output.AppendText("`nAdding Read Permissions to $user")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Reviewer"
		check_calendar
	}
	
	elseif ($calendarSetWrite.Checked -eq $True)
	{
		$output.AppendText("`r adding Read Permissions to $user")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Editor"
		check_calendar
	}
	
	elseif ($calendarSetOwner.Checked -eq $true)
	{
		$output.AppendText("`nAdding Read Permissions to $user")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Owner"
		check_calendar
	}
	
	elseif ($calendarSetRemove.Checked -eq $True)
	{
		$output.AppendText("`nRemoving permissions from $user")
		Remove-MailboxFolderPermission -identity $calendar -user $user
		$output.AppendText("`nPermissions from $user removed from $calendar")
	}
	
	else
	{
		$output.text = "No permissions set to add. Please select a level of permissions for this user."
	}
}

function terminate_user
{
	#Get User Names
	$termedUser = '{0}' -f $TerminatedUser.Text
	$fwdUser = '{0}' -f $ForwardUser.Text
	$delegUser = '{0}' -f $DelegateUser.Text
	$alias = $termedUser.Split("@")[0]
	# Clear Output Box
	$output.Text = ''
	
	# Checking for fowarding request
	If ($EnableForwardingYes.Checked -eq $True)
	{
		#Forward Email
		try
		{
			Set-Mailbox $termedUser -ForwardingAddress $fwdUser
			$output.AppendText("`nForwarding email from $termedUser to $fwdUser...")
			
			#Change Description
			Set-ADUSer $alias -Description "Terminated. Email forwarding to $fwdUser"
		}
		catch
		{
			$output.AppendText("`nError forwarding accounts. Check Terminated and Forwarding addresses.")
		}
	}
	
	# If no forwarding request, go with 
	ElseIf ($EnableForwardingNo.Checked -eq $True)
	{
		#Change Description (if not forwarding)
		$output.AppendText("done.`nChanging description...")
		Set-ADUser $alias -Description "Terminated. No Licenses"
	}
	
	#Checking for Delegate
	
	If ($DelegateAccessYes.Checked -eq $true)
	{
		Add-MailboxPermission -Identity $termedUser -User $delegUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true
	}
	
	### CHECKBOX GROUP ###
	
	If ($DisableAccountOption.Checked -eq $true)
	{
		#Disable AD Account
		$output.AppendText("`nDisabling ActiveDirectory Account...")
		Disable-ADAccount $alias
		$output.AppendText("done.")
		
		# Block Sign-In to Office 365 
		$output.AppendText("`nBlocking sign-in access to Office 365...")
		Set-MsolUser -UserPrincipalName $termedUser -BlockCredential $True
		$output.AppendText("done.")
	}
	
	If ($RemoveFromDGOption.Checked -eq $True)
	{
		#Remove from ALL Distribution Groups. 
		$output.AppendText("`nRemoving user from all Distribution Groups...")
		try
		{
			(Get-ADUser $alias -Properties MemberOf).memberOf |
			ForEach-Object {
				Remove-ADGroupMember -Identity $_ -Members $alias -Confirm $False
			}
		}
		catch
		{
			$output.AppendText("`nError. Probably failed to remove from Domain Users. Double check to be sure.")
		}
	}
	
	If ($ChangePasswordOption.Checked -eq $True)
	{
		$output.AppendText("`nGenerating randomized password and applying...")
		#Randomize Change Password 10 random characters, 5 integers, 10 more characters.
		$a = -join ((64 .. 90) + (97 .. 122) | Get-Random -Count 4 | %{ [char]$_ }) #<-- Random letters. Picks an integer between 65 and 90(lower) and 97-122(caps) and then matches that to its ASCII code in char to generate a letter. 
		$b = (10000 .. 99999) | Get-Random -Count 1 #<-- Random lumber, between 10000-99999. Only run once.
		$c = -join ((33 .. 47) | Get-Random -Count 3 | %{ [char]$_ })
		#Converts to String!
		$newpass = $a + $c + $b + $a
		$newpass = ConvertTo-SecureString -String $newpass -AsPlainText -Force
		Set-ADAccountPassword -Identity $alias -NewPassword $newpass
	}
		
	If ($RemoveLicenseOption.Checked -eq $True)
	{
		#Remove any licenses assigned, all of em.
		try
		{
			$output.appendtext("`nRemoving all Office365 Licenses...")
			(get-MsolUser -UserPrincipalName $termedUser).licenses.AccountSkuId |
			foreach{
				Set-MsolUserLicense -UserPrincipalName $termedUser -RemoveLicenses $_
			}
		}
		catch
		{
			$output.appendtext("error. Please verify that you've successfully logged into the Azure Tenant.")
		}
		$output.AppendText("done.")
	}
	
	If ($DisableASOWAOption.Checked -eq $True)
	{
		#Disable ActiveSync and OWA for Mobile Devices
		$output.AppendText("`nDisabling ActiveSync and OWA for MobileDevices...")
		Set-CASMailbox $alias -OWAEnabled $false -PopEnabled $false
		
		#Moves to Disabled User Accounts OU
		$output.AppendText("done.`nMoving to Disabled User Accounts OU...")
		$DisabledOU = "OU=Disabled User Accounts,DC=sales,DC=verrents,DC=com"
		Get-ADUser $alias | Move-AdObject -TargetPath $DisabledOU
	}
	
	#Saves a log of work done
	$output.AppendText("done.`n$termedUser has been disabled.")
	
	#Then this grabs todays date in Year-Month-Day format
	$today = Get-Date -UFormat "%Y-%m-%d"
	
	#Then we concatenate the two together
	$filename = $alias + $today
	$output.AppendText("`nSaving termination run to log folder.")
	
	<#This grabs all of the appended text from the job run and then it saves it to the Disabled User Logs folder. This folder is created when you first run the program.
	I'll add a button later to make it at your request. or better yet, add the option to save it where you want.#>
	$output.text | Out-File "\\dc1archive01\UTL\$filename.txt"
	$now = Get-Date
	Add-Content "\\dc1archive01\UTL\$filename.txt" "`nTermination carried out by $loggedInAs on $now"
	
}

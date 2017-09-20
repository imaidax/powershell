#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------
#region Control Helper Functions

function Update-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.
	
	.PARAMETER ComboBox
		The ComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ComboBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red" -Append
		Update-ComboBox $combobox1 "White" -Append
		Update-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Update-ComboBox $combobox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ComboBox]$ComboBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
}
function Update-ListBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.
	
	.PARAMETER ListBox
		The ListBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ListBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ListBox $listBox1 "Red" -Append
		Update-ListBox $listBox1 "White" -Append
		Update-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Update-ListBox $listBox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ListBox]$ListBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$listBox.Items.Clear()
	}
	
	if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])
	{
		$listBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$listBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$listBox.Items.Add($obj)
		}
		$listBox.EndUpdate()
	}
	else
	{
		$listBox.Items.Add($Items)
	}
	
	$listBox.DisplayMember = $DisplayMember
}
#endregion


$UpcomingFeaturesMsg = @"
Expanded Calendar Permissions`n
Start SharePoint Integration and Management`n
"@
$global:loggedinas = ""

#TODO: Make these loadable variables from file.

$global:MSOLAccount = ""
$global:syncServer = ""

#------------------------
# End Global Variables
#------------------------

function login{
		$username = $adminEmail.Text
		$password = $adminPassword.Text
		#These two fields grab the text in the text boxes
		#This takes the password and converts it to a secure string (Despite it already being so) and making it usable for automatic login.
		$pass = ConvertTo-SecureString -String $password -AsPlainText -Force
		$cred = New-Object -TypeName System.Management.Automation.PSCredential($username, $pass)
		try
		{
			#Connect to Office365 Tenant
			$output.AppendText("`nAttemping login as $username to Office365 Tenant..")
			$Office365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection -Name 'Office 365'
			Import-PSSession -session $Office365Session
			#Connect to Azure Tenant
			$output.AppendText("Success.`nAttempting login into Azure Active Directory..")
			Connect-MsolService -Credential $cred
			
			# Connect to On-Prem
			$output.AppendText("Success.`nAttempting login to Local Exchange Server..")
			$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$syncServer/PowerShell/ -Credential $cred -Authentication Kerberos -Name 'On Prem Session'
			Import-PSSession -session $OnPremSession
			# Succesful Login / Welcome & Titlebar Change 
			
			[console]::beep(900, 350)
			[console]::beep(1000, 350)
			[console]::beep(800, 350)
			[console]::beep(400, 350)
			[console]::beep(600, 850)
			$you = get-user -identity $username | Select-Object DisplayName
			$you = $you -replace ".*=" -replace "}"
			$output.AppendText("Success.`nLogin Succeeded.`nWelcome back $you")
			$GUI.Text = "VER Microsoft Office365 Tool. Logged in as $username"
			$global:loggedinas = "$you"
		}
		catch
		{
			$output.AppendText("`nLogin Failure")
			[console]::beep(360, 200)
			[console]::beep(300, 300)
			#Clears password after login or failure for easier reattempt
			$adminPassword.Text = ""
			return
		}
}

function get-calendarPermissions{
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

function set-calendarPermissions{
	$calendar = '{0}:\calendar' -f $CalendarEmail.Text
	$calname = Get-Mailbox
	$user = '{0}' -f $UserEmail.Text
	if ($calendarPermissions.SelectedIndex -eq 0)
	{
		$output.AppendText("`nAdding Read Permissions to $user on $calstr")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Reviewer"
		check_calendar
	}
	
	elseif ($calendarPermissions.SelectedIndex -eq 1)
	{
		$output.AppendText("`rAssigning Edit Permissions to $user on $calstr")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Editor"
		check_calendar
	}
	
	elseif ($calendarPermissions.SelectedIndex -eq 2)
	{
		$output.AppendText("`nAdding Owner Permissions to $user on $calstr")
		Add-MailboxFolderPermission -identity $calendar -user $user -AccessRights "Owner"
		check_calendar
	}
	
	elseif ($calendarPermissions.SelectedIndex -eq 3)
	{
		Remove-MailboxFolderPermission -identity $calendar -user $user
		$output.AppendText("`nPermissions from $user removed from $calstr")
	}
	
	else
	{
		$output.text = "No permissions set to add. Please select a level of permissions for this user."
	}
}

function create-user{
	Param (	[string]$name,
			[string]$fn,
			[string]$alias,
			[string]$ln,
			[string]$upn,
			[string]$ou)

		$password = generate-password("Domain")
		$revpw = generate-password("REV")
		$pw = ConvertTo-SecureString -String $password -AsPlainText -Force
	try
	{
		$output.AppendText("`nCreating user...")
		New-RemoteMailbox -name $name -DisplayName $name -FirstName $fn -LastName $ln -Password $pw -UserPrincipalName $upn -ResetPasswordOnNextLogon $true -OnPremisesOrganizationalUnit $ou -PrimarySMTPAddress $upn -Alias $alias -RemoteRoutingAddress "$alias@ver.mail.onmicrosoft.com"
		$output.AppendText("done.`nAdding e-mail aliases..")
		Set-RemoteMailbox $upn -EmailAddresses @{ add = "$upn", "$alias@verrents.com" }
		$output.AppendText("done.`n`nINFORMATION FOR MAILER FORM:`nCreated $name`nUPN Used: $upn`nDomain Password $password`nREV Password: $revpw`n`nLicense must be manually assigned.")
		$output.AppendText("`nYou may DirSync to expedite the process.")
	}
	catch
	{
		$output.AppendText("`nYou're bad.")	
	}
}

function terminate-user{
	# Get User Names
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
			$output.AppendText("`nForwarding email from $termedUser to $fwdUser... ")
			
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
		$output.AppendText("done.`nChanging description... ")
		Set-ADUser $alias -Description "Terminated. No Licenses"
	}
	
	# Checking for Delegate
	
	If ($DelegateAccessYes.Checked -eq $true)
	{
		Add-MailboxPermission -Identity $termedUser -User $delegUser -AccessRights FullAccess -InheritanceType All -AutoMapping $true
	}
	
	### CHECKBOX GROUP ###
	
	If ($DisableAccountOption.Checked -eq $true)
	{
		#Disable AD Account
		$output.AppendText("`nDisabling ActiveDirectory Account... ")
		Disable-ADAccount $alias
		$output.AppendText("done.")
		
		# Block Sign-In to Office 365 
		$output.AppendText("`nBlocking sign-in access to Office 365... ")
		Set-MsolUser -UserPrincipalName $termedUser -BlockCredential $True
		$output.AppendText("done.")
	}
	
	If ($RemoveFromDGOption.Checked -eq $True)
	{
		# Remove from ALL Distribution Groups. 
		$output.AppendText("`nRemoving user from all Distribution Groups... ")
		try
		{
			Get-ADPrincipalGroupMembership -Identity $alias | where { $_.Name -notlike "Domain Users" } | % {Remove-ADPrincipalGroupMembership -Identity $alias -MemberOf $_ -Confirm:$false}
		}
		catch
		{
			$output.AppendText("`nError occured, group removal failed. Please manually remove from groups.")
		}
	}

	if ($ChangePasswordOption.Checked -eq $True)
	{
		$output.AppendText("`nGenerating randomized password and applying... ")
		$newpass = generate-password("Strong")
		$newpass = ConvertTo-SecureString -String $newpass -AsPlainText -Force
		Set-ADAccountPassword -Identity $alias -NewPassword $newpass
	}
	if ($RemoveLicenseOption.Checked -eq $True)
	{
		# Remove any licenses assigned, all of em.
		try
		{
			$output.appendtext("`nRemoving all Office365 Licenses... ")
			(get-MsolUser -UserPrincipalName $termedUser).licenses.AccountSkuId |
			%{
				Set-MsolUserLicense -UserPrincipalName $termedUser -RemoveLicenses $_
			}
		}
		catch
		{
			$output.appendtext("error. Please verify that you've successfully logged into the Azure Tenant.")
		}
		$output.AppendText("done.")
	}
	
	if ($DisableASOWAOption.Checked -eq $True)
	{
		#Disable ActiveSync and OWA for Mobile Devices
		$output.AppendText("`nDisabling ActiveSync and OWA for MobileDevices... ")
		Set-CASMailbox -Identity $termedUser -OWAEnabled $false -PopEnabled $false -OWAforDevicesEnabled $false -ActiveSyncEnabled $false
	}
	# Hides from GAL
	
	$user = Get-ADUser $alias -Properties *
	$user.MsExchHideFromAddressLists = "True"
	Set-ADUser -Instance $user
	$output.AppendText("done.`nHiding from Global Address List... ")
	
	
	# Moves to Disabled User Accounts OU
	$output.AppendText("done.`nMoving to Disabled User Accounts OU... ")
	$DisabledOU = "OU=Disabled User Accounts,DC=sales,DC=verrents,DC=com"
	Get-ADUser $alias | Move-AdObject -TargetPath $DisabledOU
	
	# Saves a log of work done
	$output.AppendText("done.`n$termedUser has been disabled.")
	
	# Then this grabs todays date in Year-Month-Day format
	$today = Get-Date -UFormat "%Y-%m-%d"
	
	# Then we concatenate the two together
	$filename = $alias + $today
	$output.AppendText("`nSaving termination run to log folder.")
	
	<#This grabs all of the appended text from the job run and then it saves it to the Disabled User Logs folder. This folder is created when you first run the program.
	I'll add a button later to make it at your request. or better yet, add the option to save it where you want.#>
	$output.text | Out-File "\\dc1archive01\UTL\$filename.txt"
	$now = Get-Date
	Add-Content "\\dc1archive01\UTL\$filename.txt" "`nTermination carried out by $global:loggedinas on $now"
}

function get-groupdetail{
	$GroupDescription.Text = (Get-Unifiedgroup -identity $listGroups.SelectedItem).Notes
	$GroupOwner.Text = (Get-UnifiedGroup -Identity $listGroups.SelectedItem).ManagedBy
	$GMems = Get-UnifiedGroupLinks $listGroups.SelectedItem -LinkType Members | Select-Object -ExpandProperty Name
	foreach($member in $GMems){$GroupMembers.AppendText("`n$member")}
	$GroupHiddenStatus.Text = (Get-UnifiedGroup $listGroups.SelectedItem).HiddenFromAddressListsEnabled
}

function run-groupoption{
	if ($groupControlViewOption.Checked -eq $true)
	{
		get-groupdetail
	}
	
	elseif ($groupControlHideOption.Checked -eq $True)
	{
		Set-UnifiedGroup $listGroups.SelectedItem -HiddenFromAddressListsEnabled = $true
		$GroupHiddenStatus.Text = (Get-UnifiedGroup $listGroups.SelectedItem).HiddenFromAddressListsEnabled
	}
	
	elseif ($groupControlDeleteOption.Checked -eq $True)
	{
		Remove-UnifiedGroup -Identity $listGroups.SelectedItem
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show('No action selected.','Pick something!','OK','Error')
	}
}

function generate-password([string]$strong){
	If ($strong -match "Domain")
	{
		#$output.AppendText("`nGenerating randomized password and applying... ")
		$a = -join ((64 .. 90) + (97 .. 122) | Get-Random -Count 6 | %{ [char]$_ }) #<-- Random letters. Picks an integer between 65 and 90(lower) and 97-122(caps) and then matches that to its ASCII code in char to generate a letter. 
		$b = (0000 .. 9999) | Get-Random -Count 1
		$password = $a + $b
		return $password
	}
	Elseif ($strong -eq "Strong")
	{
		$password = -join ((48 .. 57) + (63) + (64 .. 90) + (33) + (35 .. 38) + (40 .. 43) + (97 .. 122) | Get-Random -Count 16 | %{ [char]$_ })
		return $password
	}
	elseif ($strong -eq "REV")
	{
		$a = (1000 .. 9999) | Get-Random -Count 1 #<-- Random lumber, between 1000-9999. Only run once.
		$b = -join ((35 .. 38) + (40 .. 43) + (63 .. 64) + 33 | Get-Random -Count 2 | %{ [char]$_ })
		$c = "Video"
		$password  = $a.ToString() + $b.ToString() + $c
		return $password
	}

	elseif ($strong -ge 1 -and $strong -le 65)
	{
		$a = -join ((48 .. 57) + (63) + (64 .. 90) + (33) + (35 .. 38) + (40 .. 43) + (97 .. 122) | Get-Random -Count $strong | %{ [char]$_ })
		return $a
	}
	
	else
	{
		$output.AppendText("`nInvalid Input.")
		return
	}
}

function Show-Inputbox{
	Param ([string]$message=$(Throw "You must enter a prompt message"),
		[string]$title = "Custom Password Length Generator",
		[string]$default)
	[reflection.assembly]::loadwithpartialname("microsoft.visualbasic") | Out-Null
	[microsoft.visualbasic.interaction]::InputBox($message, $title, $default)
}

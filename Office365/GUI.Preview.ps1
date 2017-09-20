#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-GUI_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formVERMicrosoftOffice36 = New-Object 'System.Windows.Forms.Form'
	$labelAdminEmail = New-Object 'System.Windows.Forms.Label'
	$picturebox1 = New-Object 'System.Windows.Forms.PictureBox'
	$labelAdminPassword = New-Object 'System.Windows.Forms.Label'
	$output = New-Object 'System.Windows.Forms.RichTextBox'
	$adminPassword = New-Object 'System.Windows.Forms.TextBox'
	$adminEmail = New-Object 'System.Windows.Forms.TextBox'
	$buttonLogin = New-Object 'System.Windows.Forms.Button'
	$PrimaryTabGroup = New-Object 'System.Windows.Forms.TabControl'
	$calendarTab1 = New-Object 'System.Windows.Forms.TabPage'
	$labelUserEmailAddress = New-Object 'System.Windows.Forms.Label'
	$CalendarTabInstructions = New-Object 'System.Windows.Forms.Label'
	$emailControl1 = New-Object 'System.Windows.Forms.GroupBox'
	$PresetCalPermsLabel = New-Object 'System.Windows.Forms.Label'
	$calendarPermissions = New-Object 'System.Windows.Forms.ComboBox'
	$buttonCheckPermissions = New-Object 'System.Windows.Forms.Button'
	$calEmail = New-Object 'System.Windows.Forms.Label'
	$UserEmail = New-Object 'System.Windows.Forms.TextBox'
	$CalendarEmail = New-Object 'System.Windows.Forms.TextBox'
	$buttonSetPermissions = New-Object 'System.Windows.Forms.Button'
	$TermianteUserTab = New-Object 'System.Windows.Forms.TabPage'
	$groupbox3 = New-Object 'System.Windows.Forms.GroupBox'
	$RemoveLicenseOption = New-Object 'System.Windows.Forms.CheckBox'
	$ChangePasswordOption = New-Object 'System.Windows.Forms.CheckBox'
	$DisableAccountOption = New-Object 'System.Windows.Forms.CheckBox'
	$DisableASOWAOption = New-Object 'System.Windows.Forms.CheckBox'
	$RemoveFromDGOption = New-Object 'System.Windows.Forms.CheckBox'
	$groupbox2 = New-Object 'System.Windows.Forms.GroupBox'
	$DelegateAccessNo = New-Object 'System.Windows.Forms.RadioButton'
	$DelegateAccessYes = New-Object 'System.Windows.Forms.RadioButton'
	$DelegateUser = New-Object 'System.Windows.Forms.TextBox'
	$DelegateAccessLabel = New-Object 'System.Windows.Forms.Label'
	$buttonCheckLicenses = New-Object 'System.Windows.Forms.Button'
	$buttonTerminateUser = New-Object 'System.Windows.Forms.Button'
	$ForwardUser = New-Object 'System.Windows.Forms.TextBox'
	$ForwardEmailsTo = New-Object 'System.Windows.Forms.Label'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$EnableForwardingNo = New-Object 'System.Windows.Forms.RadioButton'
	$EnableForwardingYes = New-Object 'System.Windows.Forms.RadioButton'
	$TerminatedUser = New-Object 'System.Windows.Forms.TextBox'
	$labelTerminatedUsersEmail = New-Object 'System.Windows.Forms.Label'
	$GroupControlTab = New-Object 'System.Windows.Forms.TabPage'
	$groupbox5 = New-Object 'System.Windows.Forms.GroupBox'
	$buttonRunSelection = New-Object 'System.Windows.Forms.Button'
	$GroupControlDeleteOption = New-Object 'System.Windows.Forms.RadioButton'
	$groupControlHideOption = New-Object 'System.Windows.Forms.RadioButton'
	$groupControlViewOption = New-Object 'System.Windows.Forms.RadioButton'
	$groupbox4 = New-Object 'System.Windows.Forms.GroupBox'
	$GroupOwner = New-Object 'System.Windows.Forms.Label'
	$GroupHiddenStatus = New-Object 'System.Windows.Forms.Label'
	$labelHiddenStatus = New-Object 'System.Windows.Forms.Label'
	$labelMembers = New-Object 'System.Windows.Forms.Label'
	$GroupMembers = New-Object 'System.Windows.Forms.RichTextBox'
	$GroupDescription = New-Object 'System.Windows.Forms.RichTextBox'
	$labelGroupOwner = New-Object 'System.Windows.Forms.Label'
	$labelDescription = New-Object 'System.Windows.Forms.Label'
	$listGroups = New-Object 'System.Windows.Forms.ComboBox'
	$labelOffice365Groups = New-Object 'System.Windows.Forms.Label'
	$buttonPopulateList = New-Object 'System.Windows.Forms.Button'
	$SharePointControlTab = New-Object 'System.Windows.Forms.TabPage'
	$label1 = New-Object 'System.Windows.Forms.Label'
	$UpcomingFeatures = New-Object 'System.Windows.Forms.TabPage'
	$labelExpandOffice365Group = New-Object 'System.Windows.Forms.Label'
	$menustrip1 = New-Object 'System.Windows.Forms.MenuStrip'
	$fontdialog1 = New-Object 'System.Windows.Forms.FontDialog'
	$toolstripmenuitem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$AboutHelp = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$UpcomingFeaturesHelp = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$knownBugsHelpMenu = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$oneTimeCommandsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$dirSyncMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$getLastSyncToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$formVERMicrosoftOffice36.SuspendLayout()
	$PrimaryTabGroup.SuspendLayout()
	$calendarTab1.SuspendLayout()
	$emailControl1.SuspendLayout()
	$TermianteUserTab.SuspendLayout()
	$groupbox3.SuspendLayout()
	$groupbox2.SuspendLayout()
	$groupbox1.SuspendLayout()
	$GroupControlTab.SuspendLayout()
	$groupbox4.SuspendLayout()
	$SharePointControlTab.SuspendLayout()
	$UpcomingFeatures.SuspendLayout()
	$menustrip1.SuspendLayout()
	$groupbox5.SuspendLayout()
	#
	# formVERMicrosoftOffice36
	#
	$formVERMicrosoftOffice36.Controls.Add($labelAdminEmail)
	$formVERMicrosoftOffice36.Controls.Add($picturebox1)
	$formVERMicrosoftOffice36.Controls.Add($labelAdminPassword)
	$formVERMicrosoftOffice36.Controls.Add($output)
	$formVERMicrosoftOffice36.Controls.Add($adminPassword)
	$formVERMicrosoftOffice36.Controls.Add($adminEmail)
	$formVERMicrosoftOffice36.Controls.Add($buttonLogin)
	$formVERMicrosoftOffice36.Controls.Add($PrimaryTabGroup)
	$formVERMicrosoftOffice36.Controls.Add($menustrip1)
	$formVERMicrosoftOffice36.AutoScaleDimensions = '13, 26'
	$formVERMicrosoftOffice36.AutoScaleMode = 'Font'
	$formVERMicrosoftOffice36.ClientSize = '1103, 1019'
	$formVERMicrosoftOffice36.FormBorderStyle = 'Fixed3D'
	$formVERMicrosoftOffice36.MaximizeBox = $False
	$formVERMicrosoftOffice36.Name = 'formVERMicrosoftOffice36'
	$formVERMicrosoftOffice36.Text = 'VER Microsoft Office365 Tool'
	$formVERMicrosoftOffice36.add_Load($formVERMicrosoftOffice36_Load)
	#
	# labelAdminEmail
	#
	$labelAdminEmail.AutoSize = $True
	$labelAdminEmail.Location = '454, 52'
	$labelAdminEmail.Margin = '6, 0, 6, 0'
	$labelAdminEmail.Name = 'labelAdminEmail'
	$labelAdminEmail.Size = '137, 26'
	$labelAdminEmail.TabIndex = 0
	$labelAdminEmail.Text = 'Admin Email'
	#
	# picturebox1
	#
	$picturebox1.ImageLocation = '\\dc1archive01\utl\office365.PNG'
	$picturebox1.Location = '13, 46'
	$picturebox1.Margin = '6, 6, 6, 6'
	$picturebox1.Name = 'picturebox1'
	$picturebox1.Size = '395, 94'
	$picturebox1.SizeMode = 'StretchImage'
	$picturebox1.TabIndex = 5
	$picturebox1.TabStop = $False
	#
	# labelAdminPassword
	#
	$labelAdminPassword.AutoSize = $True
	$labelAdminPassword.Location = '682, 52'
	$labelAdminPassword.Margin = '6, 0, 6, 0'
	$labelAdminPassword.Name = 'labelAdminPassword'
	$labelAdminPassword.Size = '177, 26'
	$labelAdminPassword.TabIndex = 0
	$labelAdminPassword.Text = 'Admin Password'
	#
	# output
	#
	$output.AccessibleRole = 'None'
	$output.BackColor = 'ActiveCaptionText'
	$output.Font = 'Lucida Console, 7.875pt'
	$output.ForeColor = 'LightGreen'
	$output.Location = '13, 580'
	$output.Margin = '6, 6, 6, 6'
	$output.Name = 'output'
	$output.ReadOnly = $True
	$output.ScrollBars = 'Vertical'
	$output.Size = '1083, 430'
	$output.TabIndex = 0
	$output.Text = ''
	#
	# adminPassword
	#
	$adminPassword.ImeMode = 'NoControl'
	$adminPassword.Location = '682, 84'
	$adminPassword.Margin = '6, 6, 6, 6'
	$adminPassword.Name = 'adminPassword'
	$adminPassword.PasswordChar = '●'
	$adminPassword.Size = '212, 32'
	$adminPassword.TabIndex = 2
	$adminPassword.add_KeyDown($adminPassword_KeyDown)
	#
	# adminEmail
	#
	$adminEmail.Location = '454, 84'
	$adminEmail.Margin = '6, 6, 6, 6'
	$adminEmail.Name = 'adminEmail'
	$adminEmail.Size = '212, 32'
	$adminEmail.TabIndex = 1
	$adminEmail.add_TextChanged($adminEmail_TextChanged)
	#
	# buttonLogin
	#
	$buttonLogin.Location = '906, 84'
	$buttonLogin.Margin = '6, 6, 6, 6'
	$buttonLogin.Name = 'buttonLogin'
	$buttonLogin.Size = '148, 35'
	$buttonLogin.TabIndex = 3
	$buttonLogin.Text = 'Login'
	$buttonLogin.UseVisualStyleBackColor = $True
	$buttonLogin.add_Click($buttonLogin_Click)
	#
	# PrimaryTabGroup
	#
	$PrimaryTabGroup.Controls.Add($calendarTab1)
	$PrimaryTabGroup.Controls.Add($TermianteUserTab)
	$PrimaryTabGroup.Controls.Add($GroupControlTab)
	$PrimaryTabGroup.Controls.Add($SharePointControlTab)
	$PrimaryTabGroup.Controls.Add($UpcomingFeatures)
	$PrimaryTabGroup.Location = '6, 152'
	$PrimaryTabGroup.Margin = '6, 6, 6, 6'
	$PrimaryTabGroup.Name = 'PrimaryTabGroup'
	$PrimaryTabGroup.SelectedIndex = 0
	$PrimaryTabGroup.Size = '1093, 390'
	$PrimaryTabGroup.TabIndex = 0
	#
	# calendarTab1
	#
	$calendarTab1.Controls.Add($labelUserEmailAddress)
	$calendarTab1.Controls.Add($CalendarTabInstructions)
	$calendarTab1.Controls.Add($emailControl1)
	$calendarTab1.Controls.Add($buttonCheckPermissions)
	$calendarTab1.Controls.Add($calEmail)
	$calendarTab1.Controls.Add($UserEmail)
	$calendarTab1.Controls.Add($CalendarEmail)
	$calendarTab1.Controls.Add($buttonSetPermissions)
	$calendarTab1.BackColor = 'Control'
	$calendarTab1.Location = '8, 40'
	$calendarTab1.Margin = '6, 6, 6, 6'
	$calendarTab1.Name = 'calendarTab1'
	$calendarTab1.Padding = '3, 3, 3, 3'
	$calendarTab1.Size = '1077, 342'
	$calendarTab1.TabIndex = 0
	$calendarTab1.Text = 'Calendar Permissions'
	#
	# labelUserEmailAddress
	#
	$labelUserEmailAddress.AutoSize = $True
	$labelUserEmailAddress.Location = '723, 120'
	$labelUserEmailAddress.Margin = '6, 0, 6, 0'
	$labelUserEmailAddress.Name = 'labelUserEmailAddress'
	$labelUserEmailAddress.Size = '206, 26'
	$labelUserEmailAddress.TabIndex = 9
	$labelUserEmailAddress.Text = 'User Email Address'
	#
	# CalendarTabInstructions
	#
	$CalendarTabInstructions.Location = '9, 36'
	$CalendarTabInstructions.Margin = '6, 0, 6, 0'
	$CalendarTabInstructions.Name = 'CalendarTabInstructions'
	$CalendarTabInstructions.Size = '385, 306'
	$CalendarTabInstructions.TabIndex = 0
	$CalendarTabInstructions.Text = 'Type the persons name you''re interested in adding into the Target User field and type the calendar you wish to add them to in the Calendar Address field. 

If you wish to check the permissions of an existing calendar, type the calendar in question to the Calendar Address section and click Check Permissions.'
	#
	# emailControl1
	#
	$emailControl1.Controls.Add($PresetCalPermsLabel)
	$emailControl1.Controls.Add($calendarPermissions)
	$emailControl1.Location = '393, 9'
	$emailControl1.Margin = '6, 6, 6, 6'
	$emailControl1.Name = 'emailControl1'
	$emailControl1.Padding = '6, 6, 6, 6'
	$emailControl1.Size = '300, 311'
	$emailControl1.TabIndex = 7
	$emailControl1.TabStop = $False
	$emailControl1.Text = 'Mailbox Permissions'
	#
	# PresetCalPermsLabel
	#
	$PresetCalPermsLabel.AutoSize = $True
	$PresetCalPermsLabel.Location = '12, 31'
	$PresetCalPermsLabel.Margin = '6, 0, 6, 0'
	$PresetCalPermsLabel.Name = 'PresetCalPermsLabel'
	$PresetCalPermsLabel.Size = '75, 26'
	$PresetCalPermsLabel.TabIndex = 1
	$PresetCalPermsLabel.Text = 'Preset'
	#
	# calendarPermissions
	#
	$calendarPermissions.FormattingEnabled = $True
	$calendarPermissions.Location = '12, 71'
	$calendarPermissions.Margin = '6, 6, 6, 6'
	$calendarPermissions.Name = 'calendarPermissions'
	$calendarPermissions.Size = '276, 34'
	$calendarPermissions.TabIndex = 0
	#
	# buttonCheckPermissions
	#
	$buttonCheckPermissions.Location = '723, 206'
	$buttonCheckPermissions.Margin = '6, 6, 6, 6'
	$buttonCheckPermissions.Name = 'buttonCheckPermissions'
	$buttonCheckPermissions.Size = '301, 46'
	$buttonCheckPermissions.TabIndex = 7
	$buttonCheckPermissions.Text = 'Check Permissions'
	$buttonCheckPermissions.UseVisualStyleBackColor = $True
	$buttonCheckPermissions.add_Click($buttonCheckPermissions_Click)
	#
	# calEmail
	#
	$calEmail.AutoSize = $True
	$calEmail.Location = '723, 36'
	$calEmail.Margin = '6, 0, 6, 0'
	$calEmail.Name = 'calEmail'
	$calEmail.Size = '186, 26'
	$calEmail.TabIndex = 0
	$calEmail.Text = 'Calendar Address'
	#
	# UserEmail
	#
	$UserEmail.Location = '723, 152'
	$UserEmail.Margin = '6, 6, 6, 6'
	$UserEmail.Name = 'UserEmail'
	$UserEmail.Size = '301, 32'
	$UserEmail.TabIndex = 5
	#
	# CalendarEmail
	#
	$CalendarEmail.Location = '723, 82'
	$CalendarEmail.Margin = '6, 6, 6, 6'
	$CalendarEmail.Name = 'CalendarEmail'
	$CalendarEmail.Size = '301, 32'
	$CalendarEmail.TabIndex = 6
	#
	# buttonSetPermissions
	#
	$buttonSetPermissions.Location = '723, 274'
	$buttonSetPermissions.Margin = '6, 6, 6, 6'
	$buttonSetPermissions.Name = 'buttonSetPermissions'
	$buttonSetPermissions.Size = '301, 46'
	$buttonSetPermissions.TabIndex = 8
	$buttonSetPermissions.Text = 'Set Permissions'
	$buttonSetPermissions.UseVisualStyleBackColor = $True
	$buttonSetPermissions.add_Click($buttonSetPermissions_Click)
	#
	# TermianteUserTab
	#
	$TermianteUserTab.Controls.Add($groupbox3)
	$TermianteUserTab.Controls.Add($groupbox2)
	$TermianteUserTab.Controls.Add($DelegateUser)
	$TermianteUserTab.Controls.Add($DelegateAccessLabel)
	$TermianteUserTab.Controls.Add($buttonCheckLicenses)
	$TermianteUserTab.Controls.Add($buttonTerminateUser)
	$TermianteUserTab.Controls.Add($ForwardUser)
	$TermianteUserTab.Controls.Add($ForwardEmailsTo)
	$TermianteUserTab.Controls.Add($groupbox1)
	$TermianteUserTab.Controls.Add($TerminatedUser)
	$TermianteUserTab.Controls.Add($labelTerminatedUsersEmail)
	$TermianteUserTab.BackColor = 'Control'
	$TermianteUserTab.Location = '8, 40'
	$TermianteUserTab.Margin = '6, 6, 6, 6'
	$TermianteUserTab.Name = 'TermianteUserTab'
	$TermianteUserTab.Padding = '3, 3, 3, 3'
	$TermianteUserTab.Size = '1077, 342'
	$TermianteUserTab.TabIndex = 1
	$TermianteUserTab.Text = 'Terminate User'
	#
	# groupbox3
	#
	$groupbox3.Controls.Add($RemoveLicenseOption)
	$groupbox3.Controls.Add($ChangePasswordOption)
	$groupbox3.Controls.Add($DisableAccountOption)
	$groupbox3.Controls.Add($DisableASOWAOption)
	$groupbox3.Controls.Add($RemoveFromDGOption)
	$groupbox3.Location = '50, 13'
	$groupbox3.Margin = '6, 6, 6, 6'
	$groupbox3.Name = 'groupbox3'
	$groupbox3.Padding = '6, 6, 6, 6'
	$groupbox3.Size = '285, 320'
	$groupbox3.TabIndex = 24
	$groupbox3.TabStop = $False
	$groupbox3.Text = 'Termination Options'
	#
	# RemoveLicenseOption
	#
	$RemoveLicenseOption.Location = '12, 91'
	$RemoveLicenseOption.Margin = '6, 6, 6, 6'
	$RemoveLicenseOption.Name = 'RemoveLicenseOption'
	$RemoveLicenseOption.RightToLeft = 'No'
	$RemoveLicenseOption.Size = '261, 48'
	$RemoveLicenseOption.TabIndex = 20
	$RemoveLicenseOption.Text = 'Remove Licenses'
	$RemoveLicenseOption.UseVisualStyleBackColor = $True
	#
	# ChangePasswordOption
	#
	$ChangePasswordOption.Location = '12, 142'
	$ChangePasswordOption.Margin = '6, 6, 6, 6'
	$ChangePasswordOption.Name = 'ChangePasswordOption'
	$ChangePasswordOption.RightToLeft = 'No'
	$ChangePasswordOption.Size = '261, 48'
	$ChangePasswordOption.TabIndex = 23
	$ChangePasswordOption.Text = 'Change Password'
	$ChangePasswordOption.UseVisualStyleBackColor = $True
	#
	# DisableAccountOption
	#
	$DisableAccountOption.Location = '12, 40'
	$DisableAccountOption.Margin = '6, 6, 6, 6'
	$DisableAccountOption.Name = 'DisableAccountOption'
	$DisableAccountOption.RightToLeft = 'No'
	$DisableAccountOption.Size = '238, 48'
	$DisableAccountOption.TabIndex = 19
	$DisableAccountOption.Text = 'Disable Account'
	$DisableAccountOption.UseVisualStyleBackColor = $True
	#
	# DisableASOWAOption
	#
	$DisableASOWAOption.Location = '12, 193'
	$DisableASOWAOption.Margin = '6, 6, 6, 6'
	$DisableASOWAOption.Name = 'DisableASOWAOption'
	$DisableASOWAOption.RightToLeft = 'No'
	$DisableASOWAOption.Size = '261, 48'
	$DisableASOWAOption.TabIndex = 21
	$DisableASOWAOption.Text = 'Disable AS/OWA'
	$DisableASOWAOption.UseVisualStyleBackColor = $True
	#
	# RemoveFromDGOption
	#
	$RemoveFromDGOption.Location = '12, 244'
	$RemoveFromDGOption.Margin = '6, 6, 6, 6'
	$RemoveFromDGOption.Name = 'RemoveFromDGOption'
	$RemoveFromDGOption.RightToLeft = 'No'
	$RemoveFromDGOption.Size = '256, 58'
	$RemoveFromDGOption.TabIndex = 22
	$RemoveFromDGOption.Text = 'Remove from Distribution Groups'
	$RemoveFromDGOption.UseVisualStyleBackColor = $True
	#
	# groupbox2
	#
	$groupbox2.Controls.Add($DelegateAccessNo)
	$groupbox2.Controls.Add($DelegateAccessYes)
	$groupbox2.Location = '754, 203'
	$groupbox2.Margin = '6, 6, 6, 6'
	$groupbox2.Name = 'groupbox2'
	$groupbox2.Padding = '6, 6, 6, 6'
	$groupbox2.Size = '217, 130'
	$groupbox2.TabIndex = 0
	$groupbox2.TabStop = $False
	$groupbox2.Text = 'Delegate Access?'
	#
	# DelegateAccessNo
	#
	$DelegateAccessNo.Checked = $True
	$DelegateAccessNo.Location = '12, 80'
	$DelegateAccessNo.Margin = '6, 6, 6, 6'
	$DelegateAccessNo.Name = 'DelegateAccessNo'
	$DelegateAccessNo.Size = '170, 48'
	$DelegateAccessNo.TabIndex = 12
	$DelegateAccessNo.TabStop = $True
	$DelegateAccessNo.Text = 'No'
	$DelegateAccessNo.UseVisualStyleBackColor = $True
	$DelegateAccessNo.add_CheckedChanged($DelegateAccessNo_CheckedChanged)
	#
	# DelegateAccessYes
	#
	$DelegateAccessYes.Location = '12, 28'
	$DelegateAccessYes.Margin = '6, 6, 6, 6'
	$DelegateAccessYes.Name = 'DelegateAccessYes'
	$DelegateAccessYes.Size = '170, 48'
	$DelegateAccessYes.TabIndex = 11
	$DelegateAccessYes.Text = 'Yes'
	$DelegateAccessYes.UseVisualStyleBackColor = $True
	$DelegateAccessYes.add_CheckedChanged($DelegateAccessYes_CheckedChanged)
	#
	# DelegateUser
	#
	$DelegateUser.Enabled = $False
	$DelegateUser.Location = '415, 184'
	$DelegateUser.Margin = '6, 6, 6, 6'
	$DelegateUser.Name = 'DelegateUser'
	$DelegateUser.Size = '246, 32'
	$DelegateUser.TabIndex = 6
	#
	# DelegateAccessLabel
	#
	$DelegateAccessLabel.AutoSize = $True
	$DelegateAccessLabel.Enabled = $False
	$DelegateAccessLabel.Location = '440, 153'
	$DelegateAccessLabel.Margin = '6, 0, 6, 0'
	$DelegateAccessLabel.Name = 'DelegateAccessLabel'
	$DelegateAccessLabel.Size = '206, 26'
	$DelegateAccessLabel.TabIndex = 1
	$DelegateAccessLabel.Text = 'Delegate Access To'
	#
	# buttonCheckLicenses
	#
	$buttonCheckLicenses.Location = '391, 287'
	$buttonCheckLicenses.Margin = '6, 6, 6, 6'
	$buttonCheckLicenses.Name = 'buttonCheckLicenses'
	$buttonCheckLicenses.Size = '295, 46'
	$buttonCheckLicenses.TabIndex = 8
	$buttonCheckLicenses.Text = 'Check Licenses'
	$buttonCheckLicenses.UseVisualStyleBackColor = $True
	$buttonCheckLicenses.add_Click($buttonCheckLicenses_Click)
	#
	# buttonTerminateUser
	#
	$buttonTerminateUser.Location = '391, 229'
	$buttonTerminateUser.Margin = '6, 6, 6, 6'
	$buttonTerminateUser.Name = 'buttonTerminateUser'
	$buttonTerminateUser.Size = '295, 46'
	$buttonTerminateUser.TabIndex = 7
	$buttonTerminateUser.Text = 'Terminate User'
	$buttonTerminateUser.UseVisualStyleBackColor = $True
	$buttonTerminateUser.add_Click($buttonTerminateUser_Click)
	#
	# ForwardUser
	#
	$ForwardUser.Enabled = $False
	$ForwardUser.Location = '415, 111'
	$ForwardUser.Margin = '6, 6, 6, 6'
	$ForwardUser.Name = 'ForwardUser'
	$ForwardUser.Size = '246, 32'
	$ForwardUser.TabIndex = 5
	#
	# ForwardEmailsTo
	#
	$ForwardEmailsTo.AutoSize = $True
	$ForwardEmailsTo.Location = '440, 77'
	$ForwardEmailsTo.Margin = '6, 0, 6, 0'
	$ForwardEmailsTo.Name = 'ForwardEmailsTo'
	$ForwardEmailsTo.Size = '200, 26'
	$ForwardEmailsTo.TabIndex = 5
	$ForwardEmailsTo.Text = 'Forward Emails To:'
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($EnableForwardingNo)
	$groupbox1.Controls.Add($EnableForwardingYes)
	$groupbox1.Location = '754, 13'
	$groupbox1.Margin = '6, 6, 6, 6'
	$groupbox1.Name = 'groupbox1'
	$groupbox1.Padding = '6, 6, 6, 6'
	$groupbox1.Size = '217, 130'
	$groupbox1.TabIndex = 0
	$groupbox1.TabStop = $False
	$groupbox1.Text = 'Forward Emails?'
	#
	# EnableForwardingNo
	#
	$EnableForwardingNo.Checked = $True
	$EnableForwardingNo.Location = '12, 76'
	$EnableForwardingNo.Margin = '6, 6, 6, 6'
	$EnableForwardingNo.Name = 'EnableForwardingNo'
	$EnableForwardingNo.Size = '170, 48'
	$EnableForwardingNo.TabIndex = 10
	$EnableForwardingNo.TabStop = $True
	$EnableForwardingNo.Text = 'No'
	$EnableForwardingNo.UseVisualStyleBackColor = $True
	$EnableForwardingNo.add_CheckedChanged($EnableForwardingNo_CheckedChanged)
	#
	# EnableForwardingYes
	#
	$EnableForwardingYes.Location = '12, 28'
	$EnableForwardingYes.Margin = '6, 6, 6, 6'
	$EnableForwardingYes.Name = 'EnableForwardingYes'
	$EnableForwardingYes.Size = '170, 48'
	$EnableForwardingYes.TabIndex = 9
	$EnableForwardingYes.Text = 'Yes'
	$EnableForwardingYes.UseVisualStyleBackColor = $True
	$EnableForwardingYes.add_CheckedChanged($EnableForwardingYes_CheckedChanged)
	#
	# TerminatedUser
	#
	$TerminatedUser.Location = '415, 41'
	$TerminatedUser.Margin = '6, 6, 6, 6'
	$TerminatedUser.Name = 'TerminatedUser'
	$TerminatedUser.Size = '246, 32'
	$TerminatedUser.TabIndex = 4
	#
	# labelTerminatedUsersEmail
	#
	$labelTerminatedUsersEmail.AutoSize = $True
	$labelTerminatedUsersEmail.Location = '415, 9'
	$labelTerminatedUsersEmail.Margin = '6, 0, 6, 0'
	$labelTerminatedUsersEmail.Name = 'labelTerminatedUsersEmail'
	$labelTerminatedUsersEmail.Size = '246, 26'
	$labelTerminatedUsersEmail.TabIndex = 0
	$labelTerminatedUsersEmail.Text = 'Terminated Users Email'
	#
	# GroupControlTab
	#
	$GroupControlTab.Controls.Add($groupbox5)
	$GroupControlTab.Controls.Add($groupbox4)
	$GroupControlTab.Controls.Add($listGroups)
	$GroupControlTab.Controls.Add($labelOffice365Groups)
	$GroupControlTab.Controls.Add($buttonPopulateList)
	$GroupControlTab.BackColor = 'Control'
	$GroupControlTab.Location = '8, 40'
	$GroupControlTab.Margin = '6, 6, 6, 6'
	$GroupControlTab.Name = 'GroupControlTab'
	$GroupControlTab.Padding = '3, 3, 3, 3'
	$GroupControlTab.Size = '1077, 342'
	$GroupControlTab.TabIndex = 2
	$GroupControlTab.Text = 'Group Control'
	#
	# groupbox5
	#
	$groupbox5.Controls.Add($buttonRunSelection)
	$groupbox5.Controls.Add($GroupControlDeleteOption)
	$groupbox5.Controls.Add($groupControlHideOption)
	$groupbox5.Controls.Add($groupControlViewOption)
	$groupbox5.Location = '18, 163'
	$groupbox5.Margin = '6, 6, 6, 6'
	$groupbox5.Name = 'groupbox5'
	$groupbox5.Padding = '6, 6, 6, 6'
	$groupbox5.Size = '376, 169'
	$groupbox5.TabIndex = 10
	$groupbox5.TabStop = $False
	$groupbox5.Text = 'Group Control'
	#
	# buttonRunSelection
	#
	$buttonRunSelection.Location = '224, 54'
	$buttonRunSelection.Margin = '6, 6, 6, 6'
	$buttonRunSelection.Name = 'buttonRunSelection'
	$buttonRunSelection.Size = '140, 60'
	$buttonRunSelection.TabIndex = 3
	$buttonRunSelection.Text = 'Run Selection'
	$buttonRunSelection.UseVisualStyleBackColor = $True
	$buttonRunSelection.add_Click($buttonRunSelection_Click)
	#
	# GroupControlDeleteOption
	#
	$GroupControlDeleteOption.Location = '12, 109'
	$GroupControlDeleteOption.Margin = '6, 6, 6, 6'
	$GroupControlDeleteOption.Name = 'GroupControlDeleteOption'
	$GroupControlDeleteOption.Size = '176, 34'
	$GroupControlDeleteOption.TabIndex = 2
	$GroupControlDeleteOption.TabStop = $True
	$GroupControlDeleteOption.Text = 'Delete Group'
	$GroupControlDeleteOption.UseVisualStyleBackColor = $True
	#
	# groupControlHideOption
	#
	$groupControlHideOption.Location = '12, 73'
	$groupControlHideOption.Margin = '6, 6, 6, 6'
	$groupControlHideOption.Name = 'groupControlHideOption'
	$groupControlHideOption.Size = '176, 34'
	$groupControlHideOption.TabIndex = 1
	$groupControlHideOption.TabStop = $True
	$groupControlHideOption.Text = 'Hide Group'
	$groupControlHideOption.UseVisualStyleBackColor = $True
	#
	# groupControlViewOption
	#
	$groupControlViewOption.Location = '12, 37'
	$groupControlViewOption.Margin = '6, 6, 6, 6'
	$groupControlViewOption.Name = 'groupControlViewOption'
	$groupControlViewOption.Size = '176, 34'
	$groupControlViewOption.TabIndex = 0
	$groupControlViewOption.TabStop = $True
	$groupControlViewOption.Text = 'View Details'
	$groupControlViewOption.UseVisualStyleBackColor = $True
	#
	# groupbox4
	#
	$groupbox4.Controls.Add($GroupOwner)
	$groupbox4.Controls.Add($GroupHiddenStatus)
	$groupbox4.Controls.Add($labelHiddenStatus)
	$groupbox4.Controls.Add($labelMembers)
	$groupbox4.Controls.Add($GroupMembers)
	$groupbox4.Controls.Add($GroupDescription)
	$groupbox4.Controls.Add($labelGroupOwner)
	$groupbox4.Controls.Add($labelDescription)
	$groupbox4.Location = '407, 17'
	$groupbox4.Margin = '6, 6, 6, 6'
	$groupbox4.Name = 'groupbox4'
	$groupbox4.Padding = '6, 6, 6, 6'
	$groupbox4.Size = '661, 315'
	$groupbox4.TabIndex = 9
	$groupbox4.TabStop = $False
	$groupbox4.Text = 'Group Information'
	#
	# GroupOwner
	#
	$GroupOwner.AutoSize = $True
	$GroupOwner.Location = '170, 105'
	$GroupOwner.Margin = '6, 0, 6, 0'
	$GroupOwner.Name = 'GroupOwner'
	$GroupOwner.Size = '0, 26'
	$GroupOwner.TabIndex = 0
	#
	# GroupHiddenStatus
	#
	$GroupHiddenStatus.AutoSize = $True
	$GroupHiddenStatus.Location = '596, 105'
	$GroupHiddenStatus.Margin = '6, 0, 6, 0'
	$GroupHiddenStatus.Name = 'GroupHiddenStatus'
	$GroupHiddenStatus.Size = '0, 26'
	$GroupHiddenStatus.TabIndex = 0
	#
	# labelHiddenStatus
	#
	$labelHiddenStatus.AutoSize = $True
	$labelHiddenStatus.Location = '441, 105'
	$labelHiddenStatus.Margin = '6, 0, 6, 0'
	$labelHiddenStatus.Name = 'labelHiddenStatus'
	$labelHiddenStatus.Size = '155, 26'
	$labelHiddenStatus.TabIndex = 14
	$labelHiddenStatus.Text = 'Hidden Status:'
	#
	# labelMembers
	#
	$labelMembers.AutoSize = $True
	$labelMembers.Location = '12, 146'
	$labelMembers.Margin = '6, 0, 6, 0'
	$labelMembers.Name = 'labelMembers'
	$labelMembers.Size = '109, 26'
	$labelMembers.TabIndex = 13
	$labelMembers.Text = 'Members:'
	#
	# GroupMembers
	#
	$GroupMembers.BackColor = 'Window'
	$GroupMembers.Location = '12, 178'
	$GroupMembers.Margin = '6, 6, 6, 6'
	$GroupMembers.Name = 'GroupMembers'
	$GroupMembers.ReadOnly = $True
	$GroupMembers.Size = '330, 125'
	$GroupMembers.TabIndex = 0
	$GroupMembers.Text = ''
	#
	# GroupDescription
	#
	$GroupDescription.BackColor = 'Window'
	$GroupDescription.ForeColor = 'ControlText'
	$GroupDescription.Location = '151, 30'
	$GroupDescription.Margin = '6, 6, 6, 6'
	$GroupDescription.Name = 'GroupDescription'
	$GroupDescription.ReadOnly = $True
	$GroupDescription.Size = '445, 54'
	$GroupDescription.TabIndex = 10
	$GroupDescription.Text = ''
	#
	# labelGroupOwner
	#
	$labelGroupOwner.AutoSize = $True
	$labelGroupOwner.Location = '12, 105'
	$labelGroupOwner.Margin = '6, 0, 6, 0'
	$labelGroupOwner.Name = 'labelGroupOwner'
	$labelGroupOwner.Size = '148, 26'
	$labelGroupOwner.TabIndex = 9
	$labelGroupOwner.Text = 'Group Owner:'
	#
	# labelDescription
	#
	$labelDescription.AutoSize = $True
	$labelDescription.Location = '12, 31'
	$labelDescription.Margin = '6, 0, 6, 0'
	$labelDescription.Name = 'labelDescription'
	$labelDescription.Size = '127, 26'
	$labelDescription.TabIndex = 7
	$labelDescription.Text = 'Description:'
	#
	# listGroups
	#
	$listGroups.Location = '18, 48'
	$listGroups.Margin = '6, 6, 6, 6'
	$listGroups.Name = 'listGroups'
	$listGroups.Size = '292, 34'
	$listGroups.Sorted = $True
	$listGroups.TabIndex = 6
	#
	# labelOffice365Groups
	#
	$labelOffice365Groups.AutoSize = $True
	$labelOffice365Groups.Location = '18, 18'
	$labelOffice365Groups.Margin = '6, 0, 6, 0'
	$labelOffice365Groups.Name = 'labelOffice365Groups'
	$labelOffice365Groups.Size = '188, 26'
	$labelOffice365Groups.TabIndex = 5
	$labelOffice365Groups.Text = 'Office 365 Groups'
	#
	# buttonPopulateList
	#
	$buttonPopulateList.Location = '18, 94'
	$buttonPopulateList.Margin = '6, 6, 6, 6'
	$buttonPopulateList.Name = 'buttonPopulateList'
	$buttonPopulateList.Size = '163, 38'
	$buttonPopulateList.TabIndex = 0
	$buttonPopulateList.Text = 'Populate List'
	$buttonPopulateList.UseVisualStyleBackColor = $True
	$buttonPopulateList.add_Click($buttonPopulateList_Click)
	#
	# SharePointControlTab
	#
	$SharePointControlTab.Controls.Add($label1)
	$SharePointControlTab.BackColor = 'Control'
	$SharePointControlTab.Location = '8, 40'
	$SharePointControlTab.Margin = '6, 6, 6, 6'
	$SharePointControlTab.Name = 'SharePointControlTab'
	$SharePointControlTab.Padding = '3, 3, 3, 3'
	$SharePointControlTab.Size = '1077, 342'
	$SharePointControlTab.TabIndex = 3
	$SharePointControlTab.Text = 'SharePoint Control'
	#
	# label1
	#
	$label1.AutoSize = $True
	$label1.Location = '437, 158'
	$label1.Margin = '6, 0, 6, 0'
	$label1.Name = 'label1'
	$label1.Size = '202, 26'
	$label1.TabIndex = 0
	$label1.Text = 'To Be Implemented'
	#
	# UpcomingFeatures
	#
	$UpcomingFeatures.Controls.Add($labelExpandOffice365Group)
	$UpcomingFeatures.BackColor = 'Control'
	$UpcomingFeatures.Location = '8, 40'
	$UpcomingFeatures.Margin = '6, 6, 6, 6'
	$UpcomingFeatures.Name = 'UpcomingFeatures'
	$UpcomingFeatures.Padding = '3, 3, 3, 3'
	$UpcomingFeatures.Size = '1077, 342'
	$UpcomingFeatures.TabIndex = 4
	$UpcomingFeatures.Text = 'Coming Features'
	#
	# labelExpandOffice365Group
	#
	$labelExpandOffice365Group.AutoSize = $True
	$labelExpandOffice365Group.Location = '137, 56'
	$labelExpandOffice365Group.Margin = '6, 0, 6, 0'
	$labelExpandOffice365Group.Name = 'labelExpandOffice365Group'
	$labelExpandOffice365Group.Size = '796, 182'
	$labelExpandOffice365Group.TabIndex = 0
	$labelExpandOffice365Group.Text = 'Expand Office 365 Group Control (Rudamentary group control already in)
Implement SharePoint Control, mostly focused around permissions (Ben request)
Expand on Calendar Permissions
Update Icon
"Remember me" for future logins as a checkbox option near login
Somehow implement SSO for this

'
	#
	# menustrip1
	#
	[void]$menustrip1.Items.Add($toolstripmenuitem1)
	[void]$menustrip1.Items.Add($oneTimeCommandsToolStripMenuItem)
	$menustrip1.Location = '0, 0'
	$menustrip1.Name = 'menustrip1'
	$menustrip1.Padding = '13, 4, 0, 4'
	$menustrip1.Size = '1103, 44'
	$menustrip1.TabIndex = 6
	$menustrip1.Text = 'menustrip1'
	#
	# fontdialog1
	#
	#
	# toolstripmenuitem1
	#
	[void]$toolstripmenuitem1.DropDownItems.Add($AboutHelp)
	[void]$toolstripmenuitem1.DropDownItems.Add($knownBugsHelpMenu)
	[void]$toolstripmenuitem1.DropDownItems.Add($UpcomingFeaturesHelp)
	$toolstripmenuitem1.Name = 'toolstripmenuitem1'
	$toolstripmenuitem1.Size = '76, 36'
	$toolstripmenuitem1.Text = 'Help'
	#
	# AboutHelp
	#
	$AboutHelp.Name = 'AboutHelp'
	$AboutHelp.Size = '295, 36'
	$AboutHelp.Text = 'About'
	$AboutHelp.add_Click($AboutHelp_Click)
	#
	# UpcomingFeaturesHelp
	#
	$UpcomingFeaturesHelp.Name = 'UpcomingFeaturesHelp'
	$UpcomingFeaturesHelp.Size = '295, 36'
	$UpcomingFeaturesHelp.Text = 'Upcoming Features'
	$UpcomingFeaturesHelp.add_Click($UpcomingFeaturesHelp_Click)
	#
	# knownBugsHelpMenu
	#
	$knownBugsHelpMenu.Name = 'knownBugsHelpMenu'
	$knownBugsHelpMenu.Size = '295, 36'
	$knownBugsHelpMenu.Text = 'Known Bugs'
	$knownBugsHelpMenu.add_Click($knownBugsHelpMenu_Click)
	#
	# oneTimeCommandsToolStripMenuItem
	#
	[void]$oneTimeCommandsToolStripMenuItem.DropDownItems.Add($dirSyncMenuItem)
	[void]$oneTimeCommandsToolStripMenuItem.DropDownItems.Add($getLastSyncToolStripMenuItem)
	$oneTimeCommandsToolStripMenuItem.Name = 'oneTimeCommandsToolStripMenuItem'
	$oneTimeCommandsToolStripMenuItem.Size = '215, 36'
	$oneTimeCommandsToolStripMenuItem.Text = 'Quick Commands'
	#
	# dirSyncMenuItem
	#
	$dirSyncMenuItem.Name = 'dirSyncMenuItem'
	$dirSyncMenuItem.Size = '229, 36'
	$dirSyncMenuItem.Text = 'DirSync'
	$dirSyncMenuItem.add_Click($dirSyncMenuItem_Click)
	#
	# getLastSyncToolStripMenuItem
	#
	$getLastSyncToolStripMenuItem.Name = 'getLastSyncToolStripMenuItem'
	$getLastSyncToolStripMenuItem.Size = '229, 36'
	$getLastSyncToolStripMenuItem.Text = 'Get Last Sync'
	$getLastSyncToolStripMenuItem.add_Click($getLastSyncToolStripMenuItem_Click)
	$groupbox5.ResumeLayout()
	$menustrip1.ResumeLayout()
	$UpcomingFeatures.ResumeLayout()
	$SharePointControlTab.ResumeLayout()
	$groupbox4.ResumeLayout()
	$GroupControlTab.ResumeLayout()
	$groupbox1.ResumeLayout()
	$groupbox2.ResumeLayout()
	$groupbox3.ResumeLayout()
	$TermianteUserTab.ResumeLayout()
	$emailControl1.ResumeLayout()
	$calendarTab1.ResumeLayout()
	$PrimaryTabGroup.ResumeLayout()
	$formVERMicrosoftOffice36.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formVERMicrosoftOffice36.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formVERMicrosoftOffice36.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formVERMicrosoftOffice36.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formVERMicrosoftOffice36.ShowDialog()

} #End Function

#Call the form
Show-GUI_psf | Out-Null

# Marbleous PowerShell Manager
# Being replaced by PowerShell GUI
function login()
{
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession -session $Session

}


function calendar_view()
{

    Write-Host "Please enter the email of the calendar you're interested in viewing" -ForegroundColor Yellow
    $cal = Read-Host
    Get-MailboxFolderPermission -identity ${cal}:\calendar | format-list User,AccessRights
}


function calendar_perms()
{

    function cal_add()
    {
        Write-Host "Typos Produce Red Errors. If you see red, the program didn't work, despite the success message."
        Write-host "Enter the email address of the calendar you wish to add permissions to: "
        $cal = read-Host

        Write-Host "Enter the email address of the person you wish to add to the $cal calendar: "
        $userID = Read-Host
    
        Write-Host "What level permissions do you want them to have? `n Reviewer = Read Only`n Editor = Write Access `n Owner = Owner: "
        $perms = Read-Host

        Add-MailboxFolderPermission -identity "${cal}:\calendar" -user $userID -AccessRights $perms
        Write-Host "$userID was added to $cal calendar. They were given $perms level access."
    }

    function cal_remove()
    {
        Write-Host "Typos Produce Red Errors. If you see red, the program didn't work, despite the success message."
        Write-host "Enter the email address of the calendar you wish to remove permissions from: "
        $cal = read-Host

        Write-Host "Enter the email address of the person you wish to remove from the $cal calendar: "
        $userID = Read-Host
    
        Write-Host "To remove them, you need to remove whatever permissions they currently have. What do they have now?`n Read = Reviewer `n Write = Editor `n Owner = Owner: "
        $perms = Read-Host

        Remove-MailboxFolderPermission -identity "${cal}:\calendar" -user $userID -AccessRights $perms
        Write-Host "$userID has been removed from $cal calendar. They had $perms level access."
    }

    Write-Host "0) Back to Menu`n1)Add Calendar Permissions`n2)Remove Calendar Permissons" -ForeGroundColor Yellow
    $input = read-host "Please make a choice"
    Switch($input)

    {
        '0'{menu}
        '1'{cal_add}
        '2'{cal_remove}
    }
}


function get_groups()
{

    Write-Host 'Getting groups, please wait. A window will pop up when finished.' -ForegroundColor Yellow
    $group = get-unifiedgroup -identity * | Select-Object DisplayName, Alias | Out-GridView -Title "Current Office 365 Groups. Select One to Expand"  -PassThru

     
    Get-UnifiedGroupLinks -identity $group.Alias -LinkType Members | Select-Object -Property Name  

    Write-Host "Details of the group"
    Write-Host "====================`n`n"
    Get-UnifiedGroup -identity $group.Alias | format-list Alias,Notes,ManagedBy,HiddenFromAddressListsEnabled  
}



function hide_group()
{

    $hide = Read-Host "Enter Alias of the group you wish to hide or quit to stop" 
    set-unifiedgroup $hide -hiddenfromaddresslistsenabled $true 
}


function view_group()
{

    $Group = Read-Host 'Please write the alias of the group to pull its details. Use option 4 to retreive this information'

    Write-Host "`n`nUsers in this group:"
    Write-Host "====================="
    Get-UnifiedGroupLinks -identity $Group -LinkType Members | Select-Object -Property Name  | format-list name

    Write-Host "`nDetails of the group"
    Write-Host "===================="
    Get-UnifiedGroup -identity $Group | format-list Alias,Notes,ManagedBy,HiddenFromAddressListsEnabled
}


function spo_edit()
{
    $user = Read-Host "Please enter the email of the user in question"
    $site = Read-Host "Enter the site they're from"
    $group = Read-Host "Enter the group they're in."
}


function menu()
{
    do
    {
        Write-Host "This is a basic Advanced Office365 Management Tool" -ForeGroundColor Cyan
        Write-Host "You can add to calendars, hide groups, or users.`nSometimes the program will close after pulling up a groups details. Working on workaround.`n" -ForeGroundColor Cyan
        Write-Host "0) Login" -ForeGroundColor Yellow 
        write-host "1) Change a Calendar's Permissions" -ForegroundColor Yellow
        write-host "2) List a Calendar's Permissions" -ForegroundColor Yellow
        write-host "3) Hide An Office365 Group" -ForegroundColor Yellow
        write-host "4) Get All Office 365 Groups" -ForegroundColor Yellow
        Write-Host "5) View Relevant Office365 Group Details" -ForegroundColor Yellow
        Write-Host "6) Edit A User in SharePoint Online" -ForegroundColor Yellow
        Write-Host "7) Quit, or you can just close the window" -ForegroundColor Red
        Write-Host "C) Clear the Window" -ForegroundColor Green
        $input = read-host "Please make a choice"
        Switch($input)

        {
            '0'{login}
            '1'{calendar_perms}
            '2'{calendar_view}
            '3'{hide_group}
            '4'{get_groups}
            '5'{view_group}
            '6'{spo_edit}
            '7'{"Buhbye"}
            'C'{cls}
        }
    }while($input -ne 7)
}



menu
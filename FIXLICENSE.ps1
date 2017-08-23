Set-ExecutionPolicy RemoteSigned
Import-Module MsOnline
$UserCredential = Get-Credential
Connect-MsolService -Credential $UserCredential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession -session $Session


cls

Write-host "Enter the email address of the calendar you wish to add permissions to: "

$cal = read-Host

Write-Host "Enter the email address of the person you wish to add to the $cal calendar: "

$userID = Read-Host

Write-Host "What level permissions do you want them to have? `n Read = Reviewer `n Write = Editor: "

$perms = Read-Host

if ($perms -Contains "Reviewer" -or "Editor") {

Add-MailboxFolderPermission -identity "$cal":\calendar -windowsemailaddress $userID@ver.com
}
Else{Write-Host "Incorrect format"}

Read-Host "$userID was added to $cal calendar. They were given $perms level access."



# Removes any o365 license from users account

import-module msonline 
Connect-MsolService 
#CSV file picker module start 
Function Get-FileName($initialDirectory) 
{ 
[System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | 
Out-Null 

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$OpenFileDialog.initialDirectory = $initialDirectory 
$OpenFileDialog.filter = “All files (*.*)| *.*” 
$OpenFileDialog.ShowDialog() | Out-Null 
$OpenFileDialog.filename 
} 

#CSV file picker module end 

#Variable that holds CSV file location from file picker 
$path = Get-FileName -initialDirectory “C:\” 

#CSV import command and mailbox creation loop 
import-csv $path |
	foreach {
	(get-MsolUser -UserPrincipalName $_.UserPrincipalName).licenses.AccountSkuId |
	foreach {
		Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -RemoveLicenses $_
	}
}


#Result report on licenses assigned to imported users 
import-csv $path | Get-MSOLUser | out-gridview

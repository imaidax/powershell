<# This is just a dirty remove and there's no validation. It's not even a clean script. #>


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
$path = Get-FileName -initialDirectory “c:\” 

#Window with list of available 365 licenses and their names 
$msolaccount = Get-MsolAccountSku | out-gridview -PassThru 

#Input window where you provide the license package’s name 
$AccountSKU = $msolaccount.AccountSkuId 

#CSV import command and mailbox creation loop 
import-csv $path | foreach { 
Set-MsolUser -UserPrincipalName $_.UserPrincipalName -usagelocation “US” 
Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -RemoveLicenses "ver:DESKLESSPACK"
Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -RemoveLicenses "ver:STANDARDPACK"
Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -RemoveLicenses "ver:ENTERPRISEPACK"
Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "$AccountSKU" 
} 

#Result report on licenses assigned to imported users 
import-csv $path | Get-MSOLUser | out-gridview
#Updates

#02/24/2023
#Added functionality to check if the account was created successfully. 
#At the end of the script you will see a popup box saying either it was created successfully, it was created but some attributes are missing, or it was not created successfully. 

#02/28/2023
#Added functionality to include user's initial to the Name and DisplayName fields if the two conditions below are true. Useful for when another user exists with the same first and last name.
#1. "Yes" is entered for the $includeInitial variable when prompted. 2. The length of the $initial variable is greater than 0 (not an empty string like this "").

#03/31/2023
#Added functionality to automatically set the user's password and display it when the script is finished running. 

#04/25/2023
#Added functionality to automatically populate the Master User List spreadsheet and Emp Pass spreadsheet with the required data. 

#Allows creation of popup message boxes. 
$wshell = New-Object -ComObject Wscript.Shell
[System.Windows.Forms.MessageBox]::Show("Please save and close ALL open Excel spreadsheets. Otherwise they will be forced closed later for the script to complete successfully.", "IMPORTANT", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)

$cred = Get-Credential -Username "DomainName\" -Message "Please enter Admin username and password."
#Sets initial variables to be used when running the new-ADUser and set-ADUser cmdlets below.


Write-Host "Please complete the prompts below to create new user account." -ForegroundColor Green
$num1 = Get-Random -Minimum 0 -Maximum 9
$num2 = Get-Random -Minimum 0 -Maximum 9
$num3 = Get-Random -Minimum 0 -Maximum 9
$num4 = Get-Random -Minimum 0 -Maximum 9
$pw = "Conti"+$num1+$num2+$num3+$num4+"Equans!"
$username = (Read-Host -Prompt "Enter user's GID.").ToUpper()
$firstname = Read-Host -Prompt "Enter user's first name."
$lastname = (Read-Host -Prompt "Enter user's last name.").ToUpper()
$initial = (Read-Host -Prompt "Enter user's middle initial. Leave blank if unknown.").ToUpper()
$includeInitial = (Read-Host -Prompt "Use initial in Display Name field? Answer Yes or No. This should be Yes if another user already exists with the same first and last name. You will also want to make sure you set a unique email in the next prompt.").ToUpper()
$email = Read-Host -Prompt "Enter user's email address."
$office = Read-Host -Prompt "Enter user's office location. Office location (This should be OFFICE- or ROAM- (case-sensitive and include dash) followed by the location name: EX. OFFICE-Irvine, CA)
(This designation depends on type of user. If they are they type of user that will always work in office or trailer at job site, they are OFFICE-. Any users that work that are strictly field will be ROAM-. Ex. Foreman, electricians, General Foreman, etc. Please ask manager if you have any questions or doubts.)"
$phone = Read-Host -Prompt "Enter user's phone number."  
$company = Read-Host -Prompt "Enter user's Company Name."
$department = Read-Host -Prompt "Enter user's department."
$description = Read-Host -Prompt "Enter user's job title."
$devicetype = Read-Host -Prompt "Enter user's Device Type. You can enter Laptop User, Tablet User, Email Only, or the specific Make\Model of the device."
$contractor = (Read-Host -Prompt "Is this user a contractor? Enter Yes or No. This will set the OU and EmpType for the user").ToUpper()
$procore = (Read-Host -Prompt "Does the user need Procore software? Enter Yes or No").ToUpper()
$florida = (Read-Host -Prompt "Is this user located in Florida? Answer Yes or No.").ToUpper()


#These variables may be changed based on answers to the above questions.
$emptype = "I"
$OU = "USR"
$space = " "
$fullName = $lastname+$space+$firstname


if ($contractor -eq "YES") {Set-Variable -Name "OU" -Value "EXT"}
if ($contractor -eq "YES") {Set-Variable -Name "emptype" -Value "E"}
if ($phone -eq "") {Set-Variable -Name "phone" -Value " "}
if ($department -eq "") {Set-Variable -Name "department" -Value " "}
if ($description -eq "") {Set-Variable -Name "description" -Value " "}
if ( ($initial.length -gt 0) -and ($includeInitial -eq "YES") ) {$fullName = $lastname+$space+$initial+$space+$firstname}

#Checks if Excel is open and closes it forcefully if it is.
$excelProcess = Get-Process -Name "excel" -ErrorAction SilentlyContinue

if ($excelProcess) {
    $excelProcess | Stop-Process -Force
    Write-Host "Excel process has been forcefully closed."
    Start-Sleep -Seconds 3
} else {
    Write-Host "Excel process is not running."
}


#We use try here so that if the user entered their credentials wrong, the script doesn't continue running and trying their credentials several times and locking their account out.
try {
    New-ADUser -Path "OU=$OU,OU=Accounts,OU=ORG0323,DC=D90,DC=tes,DC=local" -Name $fullName -DisplayName $fullName -GivenName $firstname -Surname $lastname -Initials $initial -UserPrincipalName $username"@equans.com" -SamAccountName $username -Accountpassword (ConvertTo-SecureString -AsPlainText $pw -Force) -ChangePasswordAtLogon $false -Enabled $true -Credential $cred -ErrorAction Stop
} catch {
    Write-Error "Script has failed. Most likely your credentials were entered incorrectly. Please try to run the script again."
    throw
}

Start-Sleep -s 5
Write-Host "Do not close window. Script running in background."
Add-ADGroupMember -Identity AllUsersGroup -Members $username -Credential $cred
Add-ADGroupMember -Identity VPNGroup -Members $username -Credential $cred
Add-ADGroupMember -Identity ProxyGroup -Members $username -Credential $cred
Add-ADGroupMember -Identity ZscalerGroup -Members $username -Credential $cred

if ($florida -eq "YES") {Add-ADGroupMember -Identity FloridaUsersGroup -Members $username -Credential $cred}
if ($procore -eq "YES") {Add-ADGroupMember -Identity ProcoreUsersGroup -Members $username -Credential $cred}
if ($company -eq "Indicon Corporation") {Add-ADGroupMember -Identity IndiconUsersGroup -Members $username -Credential $cred}

Set-ADUser -Identity $username -Description $description -EmailAddress $email -Office $office -OfficePhone $phone -Country "US" -Company $company -Division "BU09" -EmployeeID $username -EmployeeNumber $username -Department $department -replace @{"Co"='United States of America'; "CostCenter"='BU09'; "CountryCode"='840'; "EmployeeType"=$emptype; "Flags"='1'; "msDS-cloudExtensionAttribute1"='CustomValue1'; "msDS-cloudExtensionAttribute2"='CustomValue2'; "msDS-cloudExtensionAttribute3"='CustomValue3'} -Credential $cred



#Updates MasterUserList Spreadsheet

$masterworkbookPath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo1.xlsx"
$empworkbookPath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo2.xlsx"


# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Open the workbook
$masterworkbook = $excel.Workbooks.Open($masterworkbookPath)
Start-Sleep -Seconds 3
$empworkbook = $excel.Workbooks.Open($empworkbookPath)
Start-Sleep -Seconds 8
# Select the worksheet
$masterworksheet = $masterworkbook.Sheets.Item("UserList_v2")
Start-Sleep -Seconds 8
$empworksheet = $empworkbook.Sheets.Item("Sheet2")

# Get the last used row
$masterrow = $masterworksheet.UsedRange.Rows.Count + 1
$emprow = $empworksheet.UsedRange.Rows.Count + 1

# Push the variables into the worksheet
$masterworksheet.Cells.Item($masterrow, 1) = $lastname+$space+$firstname
$masterworksheet.Cells.Item($masterrow, 2) = $username
$masterworksheet.Cells.Item($masterrow, 4) = $email
$masterworksheet.Cells.Item($masterrow, 5) = $office.Split('-')[1]
$masterworksheet.Cells.Item($masterrow, 6) = $company
$masterworksheet.Cells.Item($masterrow, 7) = $devicetype
$empworksheet.Cells.Item($emprow, 1) = $firstname+$space+$lastname.Substring(0,1)+$lastname.Substring(1).toLower()
$empworksheet.Cells.Item($emprow, 2) = $pw
$empworksheet.Cells.Item($emprow, 3) = $username
$empworksheet.Cells.Item($emprow, 4) = $pw
$empworksheet.Cells.Item($emprow, 5) = $office.Split('-')[1]


# Save the workbook and close Excel
$masterworkbook.Save()
$empworkbook.Save()
Start-Sleep -Seconds 3
$excel.Quit()
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
Start-Sleep -Seconds 3


#Checks if account exists.
$checkAdUser = Get-ADUser -Filter {SamAccountName -eq $username} -Credential $cred 

#Checks if this one attribute is set. No need to check all, as the set-ADUser cmdlet will either fail and set none of the attributes, or succeed and set all.  
$checkAttributes = get-aduser $username -property * -Credential $cred | Select -ExpandProperty "msDS-cloudExtensionAttribute1" 

if ($checkAdUser -eq $null) {$wshell.Popup("User has not been created successfully. Please try again.", 0, "WARNING", 48)}

elseif ($checkAttributes -ne "CustomAttribute1") {

$wshell.Popup("User has been created, but some attributes are missing. Please add manually or rerun script.", 0, "WARNING", 48)
$excel = New-Object -ComObject Excel.Application
$wshell.Popup("User has been created successfully! The user's password is "+$pw)
$wshell.Popup("The Master User List and Emp Pass spreadsheets will now open. Please validate everything looks good.")
$masteruserlistpath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo1.xlsx"
$emppasswordpath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo2.xlsx"
$excel.Visible = $true
$openmasteruserlist = $excel.Workbooks.Open($masteruserlistpath)
$openemppasswordlist = $excel.Workbooks.Open($emppasswordpath)
Start-Sleep -Seconds 7
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
else {$wshell.Popup("User has been created successfully! The user's password is "+$pw)
$excel = New-Object -ComObject Excel.Application
$wshell.Popup("The Master User List and Emp Pass spreadsheets will now open. Please validate everything looks good.")
$masteruserlistpath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo1.xlsx"
$emppasswordpath = "$($env:USERPROFILE)\PathToSpreadsheet\userinfo2.xlsx"
$excel.Visible = $true
$openmasteruserlist = $excel.Workbooks.Open($masteruserlistpath)
$openemppasswordlist = $excel.Workbooks.Open($emppasswordpath)
Start-Sleep -Seconds 7
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}



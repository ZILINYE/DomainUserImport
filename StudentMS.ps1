Import-Module MSOnline
$username = ""   # Administrator Username
$unsecurepass = "" #Administrator Password 
$password = ConvertTo-SecureString -AsPlainText $unsecurepass -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
Connect-MsolService -Credential $cred
$StudentID = Read-Host "Please enter the Student ID 'W0XXXXXX' "
$FirstName = Read-Host "Please enter the FirstName"
$LastName = Read-Host "Please enter the LastName"
$StudentDateOfBirth = Read-Host "Please enter the Student Date Of Birth <MMDD>"
$PrincipalName = $StudentID + "<Domain Name>"
$DisplayName = $FirstName + " " + $LastName
$fourID=$StudentID.Substring($StudentID.Length -4)
$password = $fourID + $StudentDateOfBirth
Write-Host "Student ID: "+$StudentID
Write-Host "FirstName: "+$FirstName
Write-Host "LastName: "+$LastName
Write-Host "Password : "+$password
$confirm = Read-Host "All the information Correct?"
if ($confirm -eq 'y'){
    $CreateUser = New-MsolUser -DisplayName $DisplayName -FirstName $FirstName -LastName $LastName -UserPrincipalName $PrincipalName -UsageLocation CA -LicenseAssignment <LicenseCode> -Password $password -StrongPasswordRequired $False -ForceChangePassword $False
    Write-Host ("The Student account has been created")
}



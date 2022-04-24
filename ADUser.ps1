# ! Change it On Different campus (toro -- north york campus,miss -- Mississagua Campus)
$OUPath = "OU=Student,DC=toro,DC=acumen,DC=local" 
# * Initiate log user number
$CreatedUser = 0
$SkippedUser = 0
$FailedUser = 0
$CreatedOU = 0
$CreatedGP = 0
$FailedOU = 0

$logfile = @()
# * Define Log type and warning level
$Typelist = @{File = "File"; Create = "Create"; Remove = "Remove"; Move = "Move"; Set = "Set" }
$levellist = @{Warning = "Warning"; Info = "Info"; Error = "Error" }

# * Select original csv file 
function Get-File() {   
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.filter = "CSV (*.csv)|*.csv"
    [void]$FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}

# * Main function Start
function Main() {

    param (
        [Parameter(Mandatory = $true)] [string] $filepath
    )
    try {
        # Import user data from csv to powershell
        $ADUsers = @(Import-Csv -Path $filepath -ErrorAction Stop)
        Write-Host "Users imported from $filepath"
    }
    catch {
        Write-Error "Failed to read from $filepath, exiting"
        Logging -Type $Typelist.File -Level $levellist.Error -Msg $_
       
        
    }
    # For loop Each User data 
    foreach ($u in $ADUsers) {
        # Get Student's Information 
        $OriName = $u.Name 
        $Lastname = $OriName.split(",")[0] 
        $Firstname = $OriName.split(",")[1] 
        $SamAccountName = "W0$($u.EMPLID)"
        $UserPrincipalName = "$SamAccountName@$($u.CAMPUS).acumen.local"
        $DisplayName = "$Firstname $Lastname"
        $EmailAddress = $u.CAMP_EMAIL
        $Semester = $u.ACAD_PROG_PRIMARY.Substring($u.ACAD_PROG_PRIMARY.Length - 1)
        $Year = [int]$u.TERMDESC.split(' ')[1]

        # Check OU based on student's program If doesn't exist will create
        $OU, $Path = Check_OU -ProgramName $u.PROG_DESCR -Class $u.Class -Term $u.TERMDESC -Semester $Semester -Year $Year
        
        # Get group format as "C2A"
        $Group = $OU.Substring(0, 1) + $Semester + $u.Class

        # Check student's name validation in case of (1.Name already been used 2.Account already exist)
        $exist, $Name = Check_Name -Firstname $Firstname -Lastname $Lastname -Sam $SamAccountName -OU_CSV $OU -GP_CSV $Group -OU_Path $Path

        # Check if group exist
        Check_Group -GP_CSV $Group 

        $Userpassword = Get_Password -Lastname $Lastname -SamAccountName $SamAccountName -BirthDay $BirthDay

        Write-Host "Beginning With $OriName"
        
        # Check User's Name validation
        if ($exist -eq $false) {

            # Create Domain user function
            AddADUser -Name $Name `
                -GivenName $Firstname `
                -Surname $Lastname `
                -DisplayName $DisplayName `
                -SamAccountName $SamAccountName `
                -UserPrincipalName $UserPrincipalName `
                -Path $Path `
                -Password $Userpassword `
                -EmailAddress $EmailAddress
            # Add Student into their group
            Add-ADGroupMember -Identity $Group -Members $SamAccountName
            Add-ADGroupMember -Identity Students -Members $SamAccountName

        }
        else {

            # ! Opotional If waould like to set Doamin user password 
            Write-Host "The Account for $Name already exist. Skipped!"

        }
        Write-Host "End With $OriName"        
    
    }
    # Write-Host $OUGroup | Format-Table -AutoSizes
    $script:logfile | Out-File -FilePath .\Log.txt

}
function Get_Password() {
    param (
        [Parameter(Mandatory = $true)] [string] $Lastname,
        [Parameter(Mandatory = $true)] [string] $SamAccountName,
        [Parameter(Mandatory = $true)] [string] $BirthDay
    )
    $length = $SamAccountName.length
    $Textinfo = (Get-Culture).TextInfo

    $LastPrefix = $Textinfo.ToTitleCase($Lastname.Substring(0, 2).ToLower())
    $FourID = $SamAccountName.Substring($length - 4)
    $LastFourBir = $BirthDay.substring(5).Replace("-", "")
    $StuPassword = $LastPrefix + $FourID + $LastFourBir

    return $StuPassword
}

# * Set up domain user password
function Set_Password() {
    param (
        [Parameter(Mandatory = $true)] [string] $SamAccountName,
        [Parameter(Mandatory = $true)] [string] $Password
    )
    try {
        Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -ChangePasswordAtLogon $false

    }
    catch {
        Logging -Type $Typelist.Set -Level $levellist.Error -Msg "Failed to set password for $SamAccountName "
    }
}

# * Add Domain User fucntion command
function AddADUser() {
    Param
    (
        [Parameter(Mandatory = $true)] [string] $Name,
        [Parameter(Mandatory = $true)] [string] $GivenName,
        [Parameter(Mandatory = $true)] [string] $Surname,
        [Parameter(Mandatory = $true)] [string] $DisplayName,
        [Parameter(Mandatory = $true)] [string] $SamAccountName,
        [Parameter(Mandatory = $true)] [string] $UserPrincipalName,
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $Password,
        [Parameter(Mandatory = $true)] [string] $EmailAddress
    )
    try {
        New-ADUser `
            -Name $Name `
            -GivenName $GivenName `
            -Surname $Surname `
            -DisplayName $DisplayName `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UserPrincipalName `
            -Enabled $true `
            -Path $Path `
            -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force ) -ChangePasswordAtLogon $false `
            -EmailAddress $EmailAddress 
        -PasswordNeverExpired $true 
        
        Write-Host "User $Name has been created!"
        Logging -Type $Typelist.Create -Level $levellist.Info -Msg "User $Name has been created!"
        $script:CreatedUser ++
    }
    catch {

        # Todo check student's password and set it to the pattern
        Set_Password -SamAccountName $SamAccountName -Password $Password
        Logging -Type $Typelist.Create -Level $levellist.Info -Msg "User $Name has been created!"
    }

}

# * Move Student to the right OU based on Excel File
function Move_OU() {
    param (
        [Parameter(Mandatory = $true)] [string] $userDisName,
        [Parameter(Mandatory = $true)] [string] $TargetPath
    )
    try {
        Move-ADObject -Identity $userDisName -TargetPath $TargetPath
        $msg = "OU Moved $userDisName to $TargetPath"
        Write-Host $msg
        Logging -Type $Typelist.Move -Level $levellist.Warning -Msg $msg
    }
    catch {
        Logging -Type $Typelist.Move -Level $levellist.Error -Msg "Failed To Move $userDisName to $TargetPaths"
    }

}
# * Move Student to the right Group based on Excel File
function Move_Group() {
    param (
        [Parameter(Mandatory = $true)] [string] $Sam,
        [Parameter(Mandatory = $true)] [string] $OriGroup,
        [Parameter(Mandatory = $true)] [string] $NewGroup
    )
    $oldGroup = $OriGroup.Replace(" ", ",")
    
    Add-ADGroupMember -Identity $NewGroup -Members $Sam
    Add-ADGroupMember -Identity Students -Members $Sam
    try {
        Write-Host "Group Moved $Sam from $OriGroup to $NewGroup"
        Remove-ADGroupMember -Identity $oldGroup -Members $Sam -Confirm:$false 
        Logging -Type $Typelist.Remove -Level $levellist.Warning -Msg "Group Moved $Sam from $oldGroup to $NewGroup"
    
    }
    catch {
        Write-Host "user has no group membership before"
        Logging -Type $Typelist.Remove -Level $levellist.Warning -Msg $_
    }
}

# * Calculate Student's Enroll Year
function Get_EnrollYear {
    param (
        [Parameter(Mandatory = $true)] [string] $Term,
        [Parameter(Mandatory = $true)] [int] $Semester,
        [Parameter(Mandatory = $true)] [int] $Year
    )
    # * This Function for calculating student's enroll year
    # get index of the student term
    $termlist = 'Winter', 'Spring', 'Fall'
    $termindex = (0..($termlist.Count - 1)) | Where-Object { $termlist[$_] -eq $Term } 
    if ($termindex - $Semester -ge -1) {
        $yearless = 0
    }
    else {
        # Write-Host "previous year!"
        $yearless = 1 
    }
    # get student term   
    $studentTerm = $termlist[$termindex - $Semester + 1]
    # get student year
    $studentYear = $Year - $yearless

    return $studentTerm, $studentYear
    
}

# * Check and Return Student's OU based on Excel File
function Check_OU() {

    param (
        [Parameter(Mandatory = $true)] [string] $ProgramName,
        [Parameter(Mandatory = $true)] [string] $Class,
        [Parameter(Mandatory = $true)] [string] $Term,
        [Parameter(Mandatory = $true)] [int] $Semester,
        [Parameter(Mandatory = $true)] [int] $Year
    )
    <# Define OU Name list in Original Excel Data #>
    $OU = @(@{ouname = "CSTN"; programname = "Computer Sys. Technician - Net" },
        @{ouname = "Business"; programname = "Business" },
        @{ouname = "DAB"; programname = "Data Analytics for Business" },
        @{ouname = "OAHS"; programname = "Office Admin-Health Services" },
        @{ouname = "Human"; programname = "Human Resources Management" },
        @{ouname = "IBMLS"; programname = "Int. Bus. Mng-Logistics System" },
        @{ouname = "SSWG"; programname = "Social Service Worker- Geronto" })

    foreach ($o in $OU) {
        if ($ProgramName -eq $o.programname) {
            $ouprefix = $o.ouname
            break
        }
    }         

    $studentTerm, $studentYear = Get_EnrollYear -Term $Term -Semester $Semester -Year $Year
    # $FullOUName = $ouprefix+'_'+$Class+'_'+$Term.replace(' ','_')
    $FullOUName = $ouprefix + '_' + $Class + '_' + $studentTerm + '_' + $studentYear 
    $OUExist = Get-ADOrganizationalUnit -Filter { Name -eq $FullOUName }

    if ($OUExist) {
        Write-Host "Info : OU already exist. Forward to next step"
    }
    else {
        New-ADOrganizationalUnit -Name $FullOUName -Path $script:OUPath -ProtectedFromAccidentalDeletion $false
        Logging -Type $Typelist.Create -Level $levellist.Warning -Msg "OU $FullOUName has been created"
        Write-Host "Warning : OU not Found. Created OU $FullOUName" -ForegroundColor red -BackgroundColor yellow
        $script:CreatedOU += 1 
    }
    $Path = 'OU=' + $FullOUName + ',' + $script:OUPath
    # $Path = 'OU='+$FullOUName+','+$OUPath

    return $FullOUName, $Path

}
# * Check Student's Group based on excel file
function Check_Group() {
    param (
        [Parameter(Mandatory = $true)] [string] $GP_CSV
    )
    $GP = Get-ADGroup -Filter { name -eq $GP_CSV }
    if (!$GP) {
        $script:CreatedGP += 1 
        New-ADGroup -Path $script:OUPath -Name $GP_CSV -GroupCategory Security -GroupScope Global
        Logging -Type $Typelist.Create -Level $levellist.Warning -Msg "Group $GP_CSV has been created"
        Write-Host "Waning : Group not Found. Created Group $GP_CSV" -ForegroundColor red -BackgroundColor yellow
    }
}

# * Check Student's information entry function 
function Check_StudentInfo() {
    param (
        [Parameter(Mandatory = $true)] [string] $Sam,
        [Parameter(Mandatory = $true)] [string] $OU_CSV,
        [Parameter(Mandatory = $true)] [string] $GP_CSV,
        [Parameter(Mandatory = $true)] [string] $OU_Path
    )
    
    Check_Group -GP_CSV $GP_CSV
    $OnServerUser = Get-aduser -Filter { SamAccountname -eq $Sam }
    $UserDis = $OnServerUser.distinguishedName
    $OnServerGroup = (Get-ADPrincipalGroupMembership -Identity $Sam | Where-Object { $_.name.Length -eq 3 }).name
    $OnServerOU = $OnServerUser.Distinguishedname.split(",")[1].substring(3)
    
    if ($OnServerGroup -ne $GP_CSV) {
        Move_Group -Sam $Sam -OriGroup $OnServerGroup -NewGroup $GP_CSV
    }
    elseif ($OnServerOU -ne $OU_CSV) {
        Move_OU -userDisName $UserDis -TargetPath $OU_Path
        $script:ChangeUser += 1
    }
    else {
        $script:SkippedUser += 1
    }


}

# * Check Student's Name exist or already be used
function Check_Name() {
    param (
        [Parameter(Mandatory = $true)] [string] $Firstname,
        [Parameter(Mandatory = $true)] [string] $Lastname,
        [Parameter(Mandatory = $true)] [string] $Sam,
        [Parameter(Mandatory = $true)] [string] $OU_CSV,
        [Parameter(Mandatory = $true)] [string] $GP_CSV,
        [Parameter(Mandatory = $true)] [string] $OU_Path
    )
    # Check If some one using the exact same name
    Write-Host "Working on $Firstname $Lastname"
    $rename = Get-ADUser -Filter { GivenName -eq $Firstname -and Surname -eq $Lastname }
    # write-host $rename
    if ($rename) {
        $exist = $rename | Where-Object { $_.SamAccountName -eq $Sam }

        # 1 . Student Account has been breated before already (may changed OU / Semester )
        if ($exist) {
            # check Student OU and group
            Check_StudentInfo -Sam $Sam -OU_CSV $OU_CSV -GP_CSV $GP_CSV -OU_Path $OU_Path
            return $true, "$Firstname $Lastname"
             
        }
        # 2. Different Account has been created with the exact same name
        else {
            Write-Host "Warning : The Account Name has been used by others. Going to rename it.." -ForegroundColor red -BackgroundColor yellow
            $num = 0
            # check if orginal name available
            if ($rename.Name.Contains("$Firstname $Lastname")) {
                for ($num = 1; $num -ge 1; $num++) {
                    if ($rename.Name.Contains("$Firstname $Lastname $num") -eq $false) {
                        return $false, "$Firstname $Lastname $num"
                        break
                    }
                }
            }
            else {
                
                return $false, "$Firstname $Lastname"
            }
        }
    }
    # original name available
    else {
        return $false, "$Firstname $Lastname"
    }
}

# * Logging Function for recording activities
function Logging() {
    Param
    (
        [Parameter(Mandatory = $true)] [string] $Type,
        [Parameter(Mandatory = $true)] [string] $Msg,
        [Parameter(Mandatory = $true)] [string] $Level
    )
    $script:logfile += "$Level  $Type  $Msg"
}
# Get File Path
$a = Get-File

# Call function Update AD User
try {
    Main -filepath $a
}
catch {
    Write-Error "Something Wrong with opening file"
}
Write-Host "$CreatedUser Users , $CreatedOU OUs and $CreatedGP Groups has been created, Skipped $SkippedUser Users "

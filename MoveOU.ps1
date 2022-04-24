$OUPath = "OU=Student,DC=miss,DC=acumen,DC=local" 

function Get-File() {   
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.filter = "CSV (*.csv)|*.csv"
    [void]$FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}

function Main([string]$filepath) {
    try {
        $ADUsers = @(Import-Csv -Path $filepath -ErrorAction Stop)
        Write-Host "Users imported from $filepath"
    }
    catch {
        Write-Error "Failed to read from $filepath, exiting"
        Logging -Type $Typelist.File -Level $levellist.Error -Msg $_
       
        
    }
    foreach ($u in $ADUsers){
        $term = $u.TERMDESC.split(' ')[0]
        $year = [int]$u.TERMDESC.split(' ')[1]
        $SamAccountName = "W0$($u.EMPLID)"
        # $Student = Get-aduser -Filter {SamAccountname -eq $SamAccountName}
        # $DisName = $Student.distinguishedName
        $semster = [int]$u.ACAD_PROG_PRIMARY.Substring($u.ACAD_PROG_PRIMARY.Length-1)
        $OuName,$OuPath = Check_OU -ProgramName $u.PROG_DESCR -Class $u.Class -Term $term -Semester $semster -Year $year
        # MoveOU -DisName $DisName -OuPath $OuPath  
        Write-Host $OuPath

    }
}

function MoveOU([string]$DisName,[string]$OuPath){
    try {
        Move-ADObject -Identity $DisName -TargetPath $OuPath
        $msg = "OU Moved $DisName to $OuPath"
        Write-Host $msg
        # Logging -Type $Typelist.Move -Level $levellist.Warning -Msg $msg
    }
    catch {
        Write-Host $_
        # Logging -Type $Typelist.Move -Level $levellist.Error -Msg "Failed To Move $userDisName to $TargetPaths"
    }

    
}
function Check_OU([string]$ProgramName,[string]$Class,[string]$Term,[int]$Semester,[int]$Year){

    $termlist = 'Winter','Spring','Fall'
    <# Define OU Name list in Original Excel Data #>
    $OU = @(@{ouname="CSTN";programname="Computer Sys. Technician - Net"},
    @{ouname="Business";programname="Business"},
    @{ouname="DAB";programname="Data Analytics for Business"},
    @{ouname="OAHS";programname="Office Admin-Health Services"},
    @{ouname="Human";programname="Human Resources Management"},
    @{ouname="IBMLS";programname="Int. Bus. Mng-Logistics System"},
    @{ouname="SSWG";programname="Social Service Worker- Geronto"})

    foreach ($o in $OU){
                if ($ProgramName -eq $o.programname){
                    $ouprefix = $o.ouname
                    break
                }
            }

    # get index of the student term
    $termindex = (0..($termlist.Count-1)) | Where-Object {$termlist[$_] -eq $Term} 
    if ($termindex - $Semester -ge -1){
        $yearless = 0
    }else{
        # Write-Host "previous year!"
        $yearless = 1 
    }
    # get student term   
    $studentTerm = $termlist[$termindex-$Semester+1]
    # get student year
    $studentYear = $Year - $yearless

    # Write-Host "Student Year is $studentYear and Student term is $studentTerm"
    $FullOUName = $ouprefix+'_'+$Class+'_'+$studentTerm+'_'+$studentYear 
    $OUExist = Get-ADOrganizationalUnit -Filter {Name -eq $FullOUName}
    if ($OUExist){
        Write-Host "Info : OU already exist. Forward to next step"
    }else{
        ### !!! should Be changed based on Ace domain environment !!! ###
        New-ADOrganizationalUnit -Name $FullOUName -Path $script:OUPath -ProtectedFromAccidentalDeletion $false
        # Logging -Type $Typelist.Create -Level $levellist.Warning -Msg "OU $FullOUName has been created"
        Write-Host "Waning : OU not Found. Created OU $FullOUName" -ForegroundColor red -BackgroundColor yellow
        # $script:CreatedOU +=1 
    }
    $Path = 'OU='+$FullOUName+','+$script:OUPath
    # Write-Host $script:OUPath
    return $FullOUName,$Path
}

# Get File Path
$a = Get-File
# Call function Update AD User
try{
    Main -filepath $a
}
catch{
    Write-Error "Something Wrong with opening file"
}

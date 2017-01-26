# Prereqs
# Install SDS Toolkit
# https://support.office.com/en-us/article/Install-the-School-Data-Sync-Toolkit-8e27426c-8c46-416e-b0df-c29b5f3f62e1?ui=en-US&rs=en-US&ad=US&fromAR=1
# Install AZCopy
# http://xdmrelease.blob.core.windows.net/azcopy-4-2-0-preview/MicrosoftAzureStorageTools.msi
# Run PowerShell as an administrator (I use PowerShell ISE)

#####################
# Declare variables
#####################

# Work
$working_directory = 'C:\Users\smagee\Isleworth & Syon School\IT Services - Network Documentation\PowerShell\SMA\Microsoft Classroom'
$credentials_directory = 'C:\Users\smagee\Isleworth & Syon School\IT Services - Network Documentation\PowerShell\SMA\Credentials'
# Home
#$working_directory = 'C:\Users\sam_t\Google Drive\Isleworth and Syon School\Network Documentation\Office 365\PowerShell\Microsoft Classroom'
#$credentials_directory = 'C:\Users\smagee\Isleworth & Syon School\IT Services - Network Documentation\PowerShell\SMA\Credentials'

#####################
# Set $academicyear as the current academic year suffix to group names
#####################

[int]$currentyear = (get-date -uFormat %y)
[int]$currentmonth = (get-date -uFormat %m)
[int]$newschoolyearmonth = 9

if ($currentmonth -lt $newschoolyearmonth) {$academicyear = $currentyear - 1}
else {$academicyear = $currentyear}

$academicyear = ([string]$academicyear + "-" + [string]($academicyear + 1))

#####################
# Export SIMS data
#####################

#$sims_cred = (Get-Credential -Message 'Enter your sims.NET credentials').GetNetworkCredential()
# My SIMS username
$sims_username = "smagee"
# If you want to save a credentials file with your SIMS username and password encrypted inside it then run the $sim_username line above and then the line below
# Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath "$working_directory\sims_creds"
$EncryptedPasswordFile = "$credentials_directory\sims_creds"
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$sims_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sims_username,$SecureStringPassword

#convert the SecureString object to plain text using PtrToString and SecureStringToBSTR
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureStringPassword)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

& "C:\Program Files\SIMS\SIMS .net\commandReporter.exe" "/user:$($sims_creds.UserName)" "/password:$($PlainPassword)" "/report:MS Classroom - All" ("/output:$working_directory\From_SIMS_MS_Classroom_All.csv")

#####################
# Convert SIMS Student data to a list of classes/sections
#####################

[System.Collections.ArrayList]$GroupNameArray = @()
$groups = Import-Csv "$working_directory\From_SIMS_MS_Classroom_All.csv" | Select-Object class | Where-Object {$_.'class'  -ne ""} | Sort-Object Class –Unique

foreach ($group in $groups)            
{
    $GroupName = $group.'class'

    $GroupName = $GroupName -replace ('/','-')
    $GroupName = $GroupName -replace ('Class ','')
    $GroupName = $GroupName -replace (':','')
    $GroupName = $GroupName + "-" + $academicyear
    $GroupNameArray.add($(New-Object -TypeName psobject -Property @{Alias=$GroupName})) | Out-Null
}

$GroupNameArray | Export-Csv -path "$working_directory\From_PowerShell_MS_Classroom_Section_Base_Data.csv" -NoTypeInformation

#####################
# Create Section.csv
#####################

# Import base data from a csv
$simsdata = Import-Csv "$working_directory\From_PowerShell_MS_Classroom_Section_Base_Data.csv"

# Create an array to store the data you want to export into
$arraytoexport = @()

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{
        $arraytoexport +=[pscustomobject]@{
        'ID' = $item.Alias
        # Change the line below to your DfE number
        'School DfE Number' = "3134500"
        'Section Name' = $item.Alias
        }
}

#export to .csv file
$arraytoexport | Export-Csv "$working_directory\To Upload\Section.csv" -NoTypeInformation

#####################
# Create Teacher.csv
#####################

# Import base data from a csv
$simsdata = Import-Csv "$working_directory\From_SIMS_MS_Classroom_All.csv" | Select-Object 'Work Email','Staff Code' | Where-Object {$_.'Work Email' -like "*@isleworthsyon.org" -and $_.'Staff Code'  -ne ""} | Sort-Object 'Work Email','Staff Code' –Unique

# Create an array to store the data you want to export into
$arraytoexport = @()

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{
        $arraytoexport +=[pscustomobject]@{
        'ID' = $item.'Staff Code'
        # Change the line below to your DfE number
        'School DfE Number' = "3134500"
        'Username' = $item.'Work Email'
        }
}
        $arraytoexport +=[pscustomobject]@{
        # This is my ID number hard coded in so that I become a teacher in every classroom (you can't see into a classroom if your not a student or teacher)
        'ID' = "001037529512"
        # Change the line below to your DfE number
        'School DfE Number' = "3134500"
        'Username' = "smagee@isleworthsyon.org"
        }
#export to .csv file
$arraytoexport | Export-Csv "$working_directory\To Upload\Teacher.csv" -NoTypeInformation

#####################
# Create Student.csv
#####################

# Import base data from a csv
$simsdata = Import-Csv "$working_directory\From_SIMS_MS_Classroom_All.csv" | Select-Object 'Primary Email','UPN' | Where-Object {$_.'Primary Email' -like "*@isleworthsyon.org" -and $_.'UPN' -ne ""} | Sort-Object 'Primary Email' –Unique

# Create an array to store the data you want to export into
$arraytoexport = @()

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{
        $arraytoexport +=[pscustomobject]@{
        'ID' = $item.'UPN'
        # Change the line below to your DfE number
        'School DfE Number' = "3134500"
        'Username' = $item.'Primary Email'
        }
}

#export to .csv file
$arraytoexport | Export-Csv "$working_directory\To Upload\Student.csv" -NoTypeInformation

#####################
# Create StudentEnrollment.csv
#####################

# Import base data from a csv
$simsdata = Import-Csv "$working_directory\From_SIMS_MS_Classroom_All.csv" | Select-Object 'UPN','Class' | Where-Object {$_.'UPN' -ne "" -and $_.'Class' -ne ""} | Sort-Object 'UPN','Class' -Unique

# Create an array to store the data you want to export into
$arraytoexport = @()

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{

        $class = $item.'Class'
        $class = $class -replace ('/','-')
        $class = $class -replace ('Class ','')
        $class = $class -replace (':','')
        $class = $class + "-" + $academicyear
        $arraytoexport +=[pscustomobject]@{

        'Section ID' = $class
        'ID' = $item.'UPN'
        }
}

#export to .csv file
$arraytoexport | Export-Csv "$working_directory\To Upload\StudentEnrollment.csv" -NoTypeInformation

#####################
# Create TeacherRoster.csv
#####################

# Import base data from a csv
$simsdata = Import-Csv "$working_directory\From_SIMS_MS_Classroom_All.csv" | Select-Object 'Staff Code','Class' | Where-Object {$_.'Staff Code' -ne "" -and $_.'Class' -ne ""}  | Sort-Object 'Staff Code','Class' -Unique

# Create an array to store the data you want to export into
$arraytoexport = @()

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{

        $class = $item.'Class'
        $class = $class -replace ('/','-')
        $class = $class -replace ('Class ','')
        $class = $class -replace (':','')
        $class = $class + "-" + $academicyear
        $arraytoexport +=[pscustomobject]@{

        'Section ID' = $class
        'ID' = $item.'Staff Code'
        }
        
}

###########
$simsdata = Import-Csv "$working_directory\MS_Classroom_Licensed.csv" | Sort-Object 'Class' -Unique

# For each row in the imported csv build a row in the array
foreach ( $item in $simsdata )
{

        $class = $item.'Class'
        $class = $class -replace ('/','-')
        $class = $class -replace ('Class ','')
        $class = $class -replace (':','')
        $class = $class + "-" + $academicyear

$arraytoexport +=[pscustomobject]@{

        'Section ID' = $class
        # This is my ID number hard coded in so that I become a teacher in every classroom (you can't see into a classroom if your not a student or teacher)
        'ID' = "001037529512"
        }
}
###########
#export to .csv file
$arraytoexport | Export-Csv "$working_directory\To Upload\TeacherRoster.csv" -NoTypeInformation

#####################
# SDS Toolkit
#####################

# Powershell.exe -File "$working_directory\SDS ToolKit - Validation.ps1"
# Powershell.exe -File "$working_directory\SDS ToolKit - Send.ps1"
# Powershell.exe -File "$working_directory\SDS ToolKit - Compare.ps1"
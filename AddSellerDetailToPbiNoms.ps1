#-------------------------------------------------------
# Change History
#  2/28/2022 - Initial Peer Review
#  2/28/2022 - Modified to match raw PBI download schema
#  3/3/2022  - Heavily modified file read/write operations
#            - Bug flagged for duplicate header - fixed
#            - Removed dependency on header names variable
#
#--------------------------------------------------------

#--------------------------------------------------------
#
# Requirements 
# Input file must be from Nominations list and should be in .csv format
#
#---------------------------------------------------------

Import-Module AzureAD

## FUNCTION SECTION ##

Function LoginToAzure(){
    If(-not $credential){
        $credential = Connect-AzureAD
    }
}  #end function LoginToAzure


Function Get-FileName($initialDirectory){  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function Get-FileName

Function Write-Result(){
    $Outputfilename = $Inputfilename.Trim(".csv") + "_output.csv"
    $NominatorInfo | Export-csv -Path $Outputfilename -NoTypeInformation
    
} #end funtion Write-Result


## END OF FUNCTION SECTION ##

LoginToAzure

$NomHeader = 'Ceres URL','Customer TPID','Customer Name','Nominated Customer Name','Title (Friendly Name)',`
              'Created On','FTA Nomination Lifecycle Stage','State & Status Hierarchy - State','State & Status Hierarchy - Status',`
              'Subsidiary Name','ATU Group','ATU Name','Nominated By','Nominator Role','Nominator Type','Nominated Project Customer Solutions',`
              'Nominated Expected Outcome','# Nominations Starting Process Stage','Nominator Manager','Nominator Department'

$NominatorInfo = Import-Csv ($Inputfilename = Get-Filename -initialDirectory C:\temp) # -Header $NomHeader
$NominatorInfo | Add-Member -MemberType NoteProperty -Name 'Nominator Manager' -Value ''
$NominatorInfo | Add-Member -MemberType NoteProperty -Name 'Nominator Department' -Value ''

$ADnominator = ''
$manager=''
$PercentComplete = 0

Write-Output("Working on it.....")
Write-Progress -CurrentOperation "Processing Nominators" -Activity "Process" -PercentComplete $PercentComplete

$NomCount = $NominatorInfo.Count
$NomPercent = 100/$nomcount

For($nomLoop = 1;$nomLoop -le $NominatorInfo.Count;$nomLoop++){

if($NominatorInfo[$nomloop].'Ceres URL' -eq 'Total'){

    Write-Result
    Exit
    }


Write-Progress -CurrentOperation "Processing Nominators" -Activity "Lookup" -PercentComplete $PercentComplete

If($NominatorInfo[$nomLoop].'Nominated By' -ne "CustomerEmail@CustomerDomain"){

    $nomMail = $NominatorInfo[$nomLoop].'Nominated By'
    $nomName = $NominatorInfo[$nomLoop].'Nominated By' -replace ".{14}$"

    $ADnominator = Get-AzureADUser -filter "startswith(Mail, '$nomMail')" | where {$_.UserType -Like "Member"}
    $ADnominator2 = Get-AzureADUser -searchstring $nomName | where {$_.UserType -Like "Member"}
    
    if ($ADnominator){
  
        $manager = Get-AzureADUserManager -ObjectId $ADnominator.ObjectId
       
        $NominatorInfo[$nomLoop].'Nominator Manager' = $manager.DisplayName
        $NominatorInfo[$nomLoop].'Nominator Department' = $ADnominator.Department
   
    }
    ElseIf ($ADnominator2){

        if($ADnominator2.Count -ge 2){
    
            $ADnominator2 = $ADnominator2 | Where-Object {$_.CompanyName -eq 'MICROSOFT'}
        }
     
        $manager = Get-AzureADUserManager -ObjectId $ADnominator2.ObjectId
       
        $NominatorInfo[$nomLoop].'Nominator Manager' = $manager.DisplayName
        $NominatorInfo[$nomLoop].'Nominator Department' = $ADnominator2.Department
        
    }
    Else{

        $NominatorInfo[$nomLoop].'Nominator Manager' = 'Not MS' 
        $NominatorInfo[$nomLoop].'Nominator Department' = 'Not MS' 

    }

$ADnominator = ''
$ADnominator2 = ''
$manager=''
$PercentComplete += $NomPercent

    }
    }
    




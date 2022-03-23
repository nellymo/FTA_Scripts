
Import-Module AzureAD
#Connect-AzureAD

## Declarations

## The following four lines only need to be declared once in your script.
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Description."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Description."
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Description."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)

## Declarations End

## FUNCTION SECTION ##

Function LoginToAzure(){
    If(-not $credential){
        $credential = Connect-AzureAD
    }
}  #end function LoginToAzure

Function Get_Directs($person){

$directs = Get-AzureADUserDirectReport -ObjectId $person.objectid

#write-host $person.displayname

if($directs.Count -ge 1)
{
write-host $person.Displayname ' has ', $directs.Count, ' directs'

foreach($direct in $directs){

$managername = (Get-AzureADUserManager -ObjectId $direct.objectid).displayname
$sellername = $direct.displayname
$selleremail = $direct.mail

$lineToWrite = $managername + ',' + $direct.displayname + ',' + $selleremail + ',' + $direct.objectid

#write-host 'manager name', $managername

#write-host 'seller name', $sellername

Add-Content $Outputfilename $lineToWrite

$seller2 = Get_Directs -person $direct

}
}
}  #end function Get_Directs

Function Get-FileName($initialDirectory){  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog # OpenFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = “All files (*.csv)| *.csv”
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
} #end function Get-FileName

## END FUNCTION SECTION


## Main


$Inputfilename = Get-Filename -initialDirectory C:\temp
$Outputfilename = $Inputfilename #.Trim(".csv") + "_output.csv"

New-Item $Outputfilename -Force
# Writes header to the file, we add ObjectID for ease of cross referencing with the PBI nominations list of nominators. Immutable, Unique ID.
Set-Content $Outputfilename 'Manager Name, Seller Name, Seller Email, objectID'
$index=0

$top = Get-AzureADUser -SearchString "dmrug" | where {$_.UserType -Like "Member"}

$reports = Get-AzureADUserDirectReport -ObjectId $top.objectid

if($reports.Count -ge 1){

Foreach($report in $reports)

{
#write-host 'report', $report.DisplayName

$sellers = Get_Directs -person $report

#Write-Result

}
}


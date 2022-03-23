
Import-Module AzureAD
#Connect-AzureAD

## FUNCTION SECTION ##

Function LoginToAzure(){
    If(-not $credential){
        $credential = Connect-AzureAD
    }
}  #end function LoginToAzure


Function Set-FileName($initialDirectory){  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog # OpenFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = “All files (*.csv)| *.csv”
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
} #end function Get-FileName

Function Get-FileName($initialDirectory){  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function Set-FileName

## END FUNCTION SECTION

## MAIN

# Open the Processed Noms List

$NomsList = Import-Csv ($Inputfilename = Get-Filename -initialDirectory C:\temp) # -Header $NomHeader

# Open the directs list

$DirectsList = Import-Csv ($Inputfilename = Get-Filename -initialDirectory C:\temp) # -Header $NomHeader

$nomslist.Count
$DirectsList.Count

#work through the noms list and build a new result set incrementing each 'Directs' record by the nomination total

Foreach($direct in $directslist){

    #Foreach($nom in $nomslist){


    $noms = $nomslist | where {$_.'nominator objectid' -eq $direct.objectID}

    if($noms.count -ge 1){
    
        write-host $direct.'Seller Name' ,' nom count ' , $noms.count
        }
    }


param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11, 12, 13)]
    [string] $packgeID

)
Set-Location $PSScriptRoot
Clear-Host
class TagObject{
    [String] $Name
    [String] $ACTTYPE
    [String] $DATE
}

$filter = "EPCIC"+ $packgeID+".{1,}-items.csv"

$source_csv = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged" | Where-Object {$_.Name -match $filter}

$all_light_regex = Import-Csv -Delimiter ";" -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'
foreach($file in $source_csv){
    $csv = Import-Csv -Delimiter ";" -Path $file.FullName
    $output=$file.FullName.Replace(".csv","_processed.csv")

    # "Name;Description;Type;SiteName" | Out-File -FilePath $output

    $count = $csv.Count
    $result = New-Object TagObject[] $count
    # $result = New-Object TagObject[] $count
    $i = 0

    # foreach($record in $csv){
    for ($ii = 0; $ii -lt $count; $ii++){
        Write-Progress -Activity "Processing in Progress $file" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        $record = $csv[$ii]
        foreach($regex in $all_light_regex){
            if($record.Name -match $regex.Regexp){
                $tag = New-Object -TypeName TagObject
                $tag.Name = $record.Name
                $tag.ACTTYPE = $record.ACTTYPE
                $tag.DATE = $record.DATE
                $result[$ii] = $tag
                break
            }    
        }
        
    }
    Write-Output "Processign finished." 
    Write-Output "Saving file into:
    $output
    "
    
   $result | Sort-Object  -Property Name -Descending -Unique | Export-Csv -Path $output -NoTypeInformation
 
    Write-Output "File saved."
    
}
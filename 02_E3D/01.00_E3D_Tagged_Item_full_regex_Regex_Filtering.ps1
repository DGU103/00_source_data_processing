param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11, 12, 13)]
    [string] $epc

)
Set-Location $PSScriptRoot
Clear-Host
class TagObject{
    [String] $Name
    [String] $ACTTYPE
    [String] $DATE
    [string] $namingtemplate

}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Get-Job | Remove-Job

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

$Host.UI.RawUI.WindowTitle = "E3D Tagged items full Regexp check for EPC $epc"

$source_csv = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\EPCIC"+ $epc+"-E3D-Tagged-Items.csv"

$full_regexes = Import-Csv -Delimiter ";" -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv'

$inArray = Import-Csv -Delimiter ";" -Path $source_csv 
$parts = 4

[int] $partSize = [Math]::Round($inArray.count / $parts, 0)
if ($partSize -eq 0) { 
    Write-Log -Level ERROR -Message "$parts sub-arrays requested, but the input array has only $($inArray.Count) elements."
    throw }
$extraSize = $inArray.Count - $partSize * $parts
$offset = 0
$jobs_list = @()


foreach ($i in 1..$parts) {
     
    $temp = $inArray[$offset..($offset + $partSize + [bool] $extraSize - 1)]
    
    $job_id = "TAGGED_EPC" + $epc + "_Batch" + $i.ToString()



    Start-Job -Name $job_id -ScriptBlock {
        class TagObject{
            [String] $Name
            [String] $ACTTYPE
            [String] $DATE
            [string] $namingtemplate
        
        }
        $count = $args[0].count
        $full_regexes = $args[1]
        $result = New-Object  TagObject[] $count
        for ($ii = 0; $ii -lt $count; $ii++){
            $record = $args[0][$ii]

            if ($record.ACTTYPE -in ("ZONE","SITE","NOZZ")) {
                CONTINUE
            }
            foreach($regex in $full_regexes){
                if($record.Name -match ("^" + $regex.Regexp + "$")){
                    $tag = New-Object -TypeName TagObject
                    $tag.Name = $record.Name
                    $tag.ACTTYPE = $record.ACTTYPE
                    $tag.DATE = $record.DATE
                    $tag.namingtemplate = $regex.Naming_template_ID
                    $result[$ii] = $tag
                    break
                }    
            }
        
        }
        return $result} -ArgumentList $temp, $full_regexes


    $jobs_list += $job_id 
    
    $offset += $partSize + [bool] $extraSize
    if ($extraSize) { --$extraSize }
}


Wait-Job  -Name $jobs_list
# Write-host "Waiting for jobs to be completed for Package EPC$epc"

$export = $null

foreach ($job in $jobs_list) {
    $export += Receive-Job -Name $job
}

Remove-Job -Name $jobs_list

$output = $source_csv.Replace(".csv","_processed.csv")

try {

$export | Sort-Object -Property Name -Unique | Select-Object -Property NAME, ACTTYPE, DATE, NAMINGTEMPLATE | Export-Csv -Path $output -NoTypeInformation -Force -Encoding UTF8
$finished = $true
Write-Log -Level INFO -Message "TAG Export finished successfully." -finished $finished
}

catch {
    Write-Log -Level ERROR -Message "Failed to export CSV. Error: $($_.Exception.Message)"
    throw
}

#Remove-Job -Name $jobs_list
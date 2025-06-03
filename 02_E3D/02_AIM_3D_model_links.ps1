param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [int]$epc,
    # [Parameter(Mandatory=$true)]
    [ValidateSet("delete","")]
    [string]$Action
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "AIM"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

Write-Host ""

$files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged" -Filter "EPCIC$epc*_E3D-parents.csv"

#$E3D_Export = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\EPCIC" + $epc + "_E3D-parents.csv"

# if (-not (Test-Path -Path $E3D_Export)) {

#     Write-Log -Level ERROR -Message "Path $tag_path not found"
#     throw
# }

# $3d_items = Import-Csv -Path $E3D_Export -Delimiter ';'

$Tags = @{}

# Fetching data from MTR

$path = "\\Qamv3-sapp243\gdp\GDP_StagingArea\MP\MTR\MTR_EPCIC" + $epc + "_Tag_Load_to_AIM-A.csv"

if (-not (Test-Path -Path $path)) {

    Write-Log -Level ERROR -Message "Path $path not found"
    throw
}

$temp = Import-Csv -Path $path  -Delimiter ","
foreach($record in $temp){
    if (-not($Tags[$record.Tag_Number])) {
        $Tags.Add($record.Tag_Number , $record)
    }
}




foreach ($file in $files) {

    $clean = $file.Name -replace ('_E3D-parents.csv', '')

    $3d_items = Import-Csv -Path $file.FullName -Delimiter ';'

$inArray = $3d_items
$parts = 4

[int] $partSize = [Math]::Round($inArray.count / $parts, 0)
if ($partSize -eq 0) { throw "$parts sub-arrays requested, but the input array has only $($inArray.Count) elements." }
$extraSize = $inArray.Count - $partSize * $parts
$offset = 0
$jobs_list = @()

    foreach ($i in 1..$parts) {
        $temp = $inArray[$offset..($offset + $partSize + [bool] $extraSize - 1)]
        $job_id = "3D_LINKS_EPC" + $epc + "_Batch" + $i.ToString()
    
        Start-Job -Name $job_id -ScriptBlock {

            $3d_items = $args[0]
            $Tags = $args[1]
            $Action = $args[2]
            class ModelLink {
                [string] $Model
                [String] $Tag_Number
                [String] $Ref3D
                [string] $Platform
                [string] $Action
            }
            # $count = $3d_items
            $result = @()
            
            
            foreach($3d_item in $3d_items){
                # Write-Progress -Activity "Processing in Progress $file" -Status "$i% Complete: from $count" -PercentComplete $i
                # $i = $i+ 100 / $count
                $record = New-Object ModelLink
                $record.Action = $Action
                <# If the 3D item found in Tag numbers, then we creating a simple link to 3D geometry #>
                if ($Tags[$3d_item.NAME.Substring(1)] ) {
                    $record.model = $3d_item.NAME.Substring(1,4) + '_3D_MODEL'
                    $record.Platform = $3d_item.NAME.Substring(1,4)
                    $record.Tag_Number = $3d_item.NAME.Substring(1)
                    $record.Ref3D = $3d_item.NAME

                    $result += $record
                    CONTINUE
                }
					
                <# 
                    If 3D item is not a Tag, then we need to check all parents, 
                    and if one of the parents can be found in list of Tags
                    then we building relevant link 
                #>
                <# 
                    If no named child items in E3D hier then skip
                #>
                if (-not($3d_item.parents)) { 
                    # Write-Host $3d_item -ForegroundColor DarkRed
                    CONTINUE }
                <# 
                    If there are named child items in E3D hier then all of them shall refer to
                    the parent Tagged or Non-Tagged item
                #>

                # Child elements extraction

                foreach ($parent in $3d_item.parents.split('#')) {
                    if ($3d_item.NAME.Length -lt 5) {continue}
                    if ($Tags[$parent]) {
                        $record.model = $parent.Substring(0,4) + '_3D_MODEL'
                        $record.Platform = $parent.Substring(0,4)

                        if (($epc -eq '11') -and ($parent.contains('_'))) {
                            $record.Tag_Number = $parent.split('_')[0]
                        }

                       else {$record.Tag_Number = $parent}
                        $record.Ref3D = $3d_item.NAME
                        $result += $record
                       
                    }
                }
                
                
            }


            return $result
        } -ArgumentList $temp, $Tags, $action

        $jobs_list += $job_id 
        $offset += $partSize + [bool] $extraSize
        if ($extraSize) { --$extraSize }
    }

    
    Wait-Job  -Name $jobs_list

$job_results = @()
Write-Log -Level INFO -Message "Recieving Jobs..."
foreach ($job in $jobs_list) {
    $job_results += Receive-Job -Name $job
}

Remove-Job -Name $jobs_list

$export = @()
$test_array = @{}
Write-Log -Level INFO -Message "Selecting results..."
foreach ($item in $job_results) {
    if (-NOT($test_array[$item.ref3D])) {
  
    $test_array.Add($item.ref3D,$item)
    $export += $item
    }
}


#Final Export

$root_path = "\\qamv3-sapp243\GDP\GDP_StagingArea\MP\ref3D\"
#$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\3DModel\"

$export_file_path = $root_path + $clean + "-MTR_3D_items2Tag_refs.csv"

Write-Log -Level INFO -Message "Exporting results..."

try {
$export | Select-Object -Property Model,Tag_Number,Ref3D,Platform,Action | Export-Csv -Path $export_file_path  -NoTypeInformation -Encoding UTF8 -Force
$finished = $true
Write-Log -Level INFO -Message "Links Export finished successfully" -finished $finished

}

catch {
    Write-Log -Level ERROR -Message "Failed to export CSV. Error: $($_.Exception.Message)"
    throw
}

}


Remove-Job -Name $jobs_list
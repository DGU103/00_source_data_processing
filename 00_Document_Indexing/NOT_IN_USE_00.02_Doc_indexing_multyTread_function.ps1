function RunIndexing {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet('05','06','11','12','13')]
        [String]$epc
    )
    if($epc -in @(11,12,13)){$root_path = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\"}
    elseif($epc -eq '06'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\"}
    elseif($epc -eq '05'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\"}

    $log_path = $root_path + "logs\Parallel_tread_EPC_" + $epc + ".log"
    Start-Transcript -Path $log_path

    $files_Dir = $root_path + "EPC" + $epc + "_Source"
    $tag_report = $root_path + "EPCIC"+  [string]$epc +"_indexing_report.csv"
    if(Test-Path -Path $tag_report){Remove-Item -Path $tag_report}
    $parts = 4
    $inArray = Get-ChildItem -Path $files_Dir -Filter *.pdf -Recurse
    [int] $partSize = [Math]::Round($inArray.count / $parts, 0)
    if ($partSize -eq 0) { throw "$parts sub-arrays requested, but the input array has only $($inArray.Count) elements." }
    $extraSize = $inArray.Count - $partSize * $parts
    $offset = 0
    $jobs_list = @()

    foreach ($i in 1..$parts) {
        $temp = @()
        foreach ($currentItemName in $inArray[$offset..($offset + $partSize + [bool] $extraSize - 1)]) {
            $temp  += $currentItemName}
        $offset += $partSize + [bool] $extraSize
        if ($extraSize) { --$extraSize }
        $batch_file = $root_path + "Temp\EPC" + $epc + "_batch" + $i + ".csv"
        $temp | Select-Object -Property FullName | Export-Csv -Path $batch_file -NoTypeInformation
        # Start-Process -FilePath "powershell.exe" -ArgumentList "-File `".\00_Document_Indexing\00_Document_Indexing.ps1`" -epc `"$epc`" -batch `"$i`" -batch_file `"$batch_file`" " 

        $job_id = "EPC" + $epc + "_Batch" + $i.ToString()
        Start-Job -Name $job_id -FilePath ".\00_Document_Indexing\00_Document_Indexing.ps1" -ArgumentList $epc,$i,$batch_file
        $jobs_list += $job_id 
    }
    # $state = $true
    # $jobs = get-job
    # while($state){
    #     foreach($job in $jobs){
    #         if($job.State -eq 'Running'){
    #             Write-Host $job.ChildJobs[0].Progress
    #         }
    #         else{$state = $false}
    #     }
    # }
    
    Wait-Job  -Name $jobs_list
    Remove-Job -Name $jobs_list
    Stop-Transcript 
}


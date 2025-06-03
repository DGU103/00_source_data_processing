
Set-Location $PSScriptRoot

#Remove-Item -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Non-Tagged" -Include *_validation_report.csv  -Recurse -Exclude *Archive*

$dir_files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Non-Tagged" -Filter *.csv #| where{$_.BaseName -match 'EPCIC13.{1,}MCT'}


ForEach($file in $dir_files)
{
#$file_path = $_.FullName

#Copy-Item -Path $file.FullName -Destination $file.Directory.FullName.Replace('Source','Result')
$result_csv_path = $file.FullName.Replace('.csv','_validation_report.csv')

$source_csv = Get-Content -Path $file.FullName

$headers_source = ""

if($source_csv.GetType().Name -eq "String"){$headers_source = $source_csv}else{$headers_source = $source_csv.Item(0)}

$headers_target = ""
foreach($header in $headers_source.Split(';')){
    $headers_target = $headers_target + ";"+ $header + "_validation"
}
$headers_source = $headers_source + $headers_target

if($source_csv.GetType().Name -eq "String"){$source_csv = $headers_source}else{$source_csv.Item(0) = $headers_source}



$source_csv | Set-Content $result_csv_path

if($null -ne $csv){Clear-Variable -Name csv -Force}
$csv = $null
if($file.Name -match "-Catalog-"){
    
   $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
   $catref_regex = "\/(TEP|MCD|HHK|QMC|LTM)\/[A-Z0-9]-(0000|[A-Z]{3}[0-9]|GRE)\/(ATTA|BEND|FBLI|CAP-|LISE|COUP|CROS|ELBO|FILT|FTUB|FLAN|GASK|INST|NOZZ|OLET|TUBE|PCOM|REDU|TEE-|UNIO|VALV|WELD)[A-Z0-9]{3}[0A-Z]{2}000\/[A-Z0-9]{3}\/[0-9]{2}\/([A-B][A-Z]|00){3}"
   $spco_name_regexp = "\/(TEP|MCD|HHK|QMC|LTM)_TMP_([A-Z]{2}[0-9]|((S|I)[A-Z]{3})\/[A-Z0-9])-(0000|[A-Z]{3}[0-9]|GRE)\/(ATTA|BEND|FBLI|CAP-|LISE|COUP|CROS|ELBO|FILT|FTUB|FLAN|GASK|INST|NOZZ|OLET|TUBE|PCOM|REDU|TEE-|UNIO|VALV|WELD)[A-Z0-9]{3}[0A-Z]{2}000\/[A-Z0-9]{3}\/[0-9]{2}\/([A-B][A-Z]|00){3}"
   $MATXT_name_regexp = "\/(TEP|MCD|HHK|QMC|LTM)\/[A-Z0-9_]{1,}$"
   $cate_name_regex = "\/(TEP|MCD|HHK|QMC|LTM)\/[A-Z0-9]-(0000|[A-Z]{3}[0-9]|GRE)\/(ATTA|BEND|FBLI|CAP-|LISE|COUP|CROS|ELBO|FILT|FTUB|FLAN|GASK|INST|NOZZ|OLET|TUBE|PCOM|REDU|TEE-|UNIO|VALV|WELD)[A-Z0-9]{3}[0A-Z]{2}000\/[A-Z0-9]{3}"
   $sdte_name_regex = "\/(TEP|MCD|HHK|QMC|LTM)\/[A-Z0-9]-(0000|[A-Z]{3}[0-9]|GRE)\/(ATTA|BEND|FBLI|CAP-|LISE|COUP|CROS|ELBO|FILT|FTUB|FLAN|GASK|INST|NOZZ|OLET|TUBE|PCOM|REDU|TEE-|UNIO|VALV|WELD)[A-Z0-9]{3}[0A-Z]{2}000\/[A-Z0-9]{3}/[0-9]{2}\/SDTE"
   $CMPREF_name_regexp = "\/(TEP|MCD|HHK|QMC|LTM)\/[A-Z0-9]-(0000|[A-Z]{3}[0-9]|GRE)\/(ATTA|BEND|FBLI|CAP-|LISE|COUP|CROS|ELBO|FILT|FTUB|FLAN|GASK|INST|NOZZ|OLET|TUBE|PCOM|REDU|TEE-|UNIO|VALV|WELD)[A-Z0-9]{3}[0A-Z]{2}000\/[A-Z0-9]{3}\/[0-9]{2}\/([A-B][A-Z]|00){3}-W\/.{1,}$"
   
   $count = $csv.Count
   $i = 0
   foreach($record in $csv){
        Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
       $i = $i+ 100 / $count
  
       if($record.SPCO -match $spco_name_regexp){$record.SPCO_validation = "VALID"} else{$record.SPCO_validation = "NOT VALID"}
       if($record.CATREF -match $catref_regex){$record.CATREF_validation = "VALID"} else{$record.CATREF_validation = "NOT VALID"}
       if($record.CATE -match $cate_name_regex){$record.CATE_validation = "VALID"} else{$record.CATE_validation =  "NOT VALID"}
       if($record.SDTE -match $sdte_name_regex){$record.SDTE_validation = "VALID"} else{$record.SDTE_validation = "NOT VALID"}
       if($record.MATXT -match $MATXT_name_regexp){$record.MATXT_validation = "VALID"} else{$record.MATXT_validation = "NOT VALID"}
       if($record.CMPREF -match $CMPREF_name_regexp){$record.CMPREF_validation = "VALID"} else{$record.CMPREF_validation = "NOT VALID"}
   }
}
elseif($file.Name -match "-ELE-MCT-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    # Attachment 4 � MCT Naming #
    $MCT_regexp = "\/AS[A-Z]{3}-MCT-(YY|[0-9]{2})[0-9]{4}(H|E|N|C)"

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count

        if($record.Name -match $MCT_regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    
    }
}
elseif($file.Name -match "-INS-MCT-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    # Attachment 4 � MCT Naming #
    $MCT_regexp = "\/AS[A-Z]{3}-MCT-(YY|[0-9]{2})[0-9]{4}(I|T|P|V|A|U)"

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count

        if($record.Name -match $MCT_regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    
    }
}

elseif($file.Name -match "-ELE-Tray-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-[0-9]{3,4}-(NLV|ELV|HV)-(LD|TR)-[0-9]{4}"
    # Attachment 2 � Electrical Cable Ladder / Tray Naming: #
    #.\EPCIC11-ELE-Tray-Naming.csv    
    
    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-ELE-Support-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-ELS-(AL[1-9]|BL[1-9]|BR|FL)-[0-9]{4}(_H)?"
    # Attachment 3 - Electrical Support Naming #
    #.\EPCIC11-ELE-Support-Naming.csv
    
    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-ELE-earthing-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-EPE-(YY|[0-9]{2})[0-9]{2}" 
    # Attachment 5 � Earth Bar Naming #

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-INS-Tray-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-[0-9]{3,4}-IT(01|03|04|00)-(LD|TR)-[0-9]{4}"
    # Attachment 6 � Instrumentation Cable Ladder and Tray Naming #

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-TEL-Tray-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-[0-9]{3,4}-(TEL|PGA|PGB)-(LD|TR)-[0-9]{4}"
    # Attachment 7 � Telecom Cable Ladder and Tray Naming #

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-Pipe-Support-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{3}-(PSS|PST|SPS|PSP)-(AL[1-9]|BL[1-9]|BR|FL)-[0-9]{4}$"
    # Attachment 9 � Pipe Support Naming #

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
elseif($file.Name -match "-Trim-"){
    $csv = Import-Csv -Delimiter ";" -Path $result_csv_path
    $regexp =  "\/AS[A-Z]{2}-([0-9]{1,4}""|[0-9]{1,2}\.[0-9]{1,2}"")-TR-[A-Z]{1,2}[0-9]{4}[A-Z][0-9][A-Z]?-[A-Z]{2}[0-9]"
    # Attachment 9 � Pipe Support Naming #

    $count = $csv.Count
    $i = 0
    foreach($record in $csv){
         Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
        $i = $i+ 100 / $count
        if($record.Name -match $regexp){$record.Name_validation = "VALID"} else{$record.Name_validation = "NOT VALID"}
    }
}
if($null -ne $csv){$csv | Export-Csv -Path $result_csv_path  -NoTypeInformation -Force}
#else{Set-Content -Value "ERROR FNF" -Path $result_csv_path.Replace(".csv","_FNF.csv")}

}

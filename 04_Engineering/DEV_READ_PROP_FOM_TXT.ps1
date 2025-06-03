
$content = Get-Content -Path W:\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\PROPs\EPC13_Property_report_part_6000.txt

for ($i = 0; $i -lt $content.Count; $i++) {
    if ($content[$i] -match 'Attributes') {
        $Tag_number= $content[$i+1].split()[1]
    }
    if ($content[$i+1].split()[1] -match 'unset') {
        CONTINUE
    }
    $attname = $content[$i+1].split()[0]
    $attval = $content[$i+1].split()[1]
    $val = $Tag_number + ';' + $attname + ';' + $attval
    $val
}
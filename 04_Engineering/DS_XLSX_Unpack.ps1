$path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\DS_Templates\EPCIC-12"

$files = Get-ChildItem -Filter '*.xlsx' -Path $path
foreach ($file in $files)
{
    Rename-Item $file.FullName -NewName $file.FullName.Replace('.xlsx', '.zip')
    
}



$files = Get-ChildItem -Filter '*.zip' -Path $path
foreach ($file in $files)
{
   New-Item -Path $file.FullName.Replace('.zip','') -ItemType Directory
    
}

$files = Get-ChildItem -Filter '*.zip' -Path $path
foreach ($file in $files)
{
    $destpath = $file.FullName.Replace('.zip','')
   Expand-Archive -Path $file.FullName -DestinationPath $destpath
    
}

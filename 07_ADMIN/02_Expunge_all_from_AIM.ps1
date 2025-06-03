
$argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\Expunge.pmlmac'
Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\02_E3D\macro\Expunge.pmlmac
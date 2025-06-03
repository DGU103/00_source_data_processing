param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)


$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "EI"

. "$PSScriptRoot\..\Common_Functions.ps1"

#Forcing custom_evars.bat

Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

## EPC 11/13  Electrical ##

if ($epc -eq '13') {

  $finished = $false

  try {

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT [TagNo] as TagID
	,[PropName] as AttributeID
    ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData 
  ON [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList].[InstKEY] = [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData.[ObjectId]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef]
  ON [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData.[PropertyDefID] = [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef].[PropertyDefID]
  WHERE [TagNo] like 'ASBJ%'

UNION

SELECT [EquipmentNo] as TagID
	  ,[PropName] as AttributeID
      ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData] 
  ON [RYA-EI-TE-XXX-MCD].[dbo].[Equipment].[EquipId] = [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[ObjectId]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef]
  ON [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[PropertyDefID] = [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef].[PropertyDefID]
  WHERE [EquipmentNo]  like 'ASBJ%'

UNION
  
  
SELECT [TagNo] as TagID
      ,[PropName] as AttributeID
      ,[PropVal] as AttributeValue
	  ,'' as UnitOfMeasureID	
FROM (
SELECT      [TagNo]
      ,CAST([ISAFunc]               AS varchar) as [ISAFunc]             
      ,CAST([LNumber]               AS varchar) as [LNumber]             
      ,CAST([Suffix]                AS varchar) as [Suffix]              
      ,CAST([Prefix]                AS varchar) as [Prefix]              
      ,CAST([Description]           AS varchar) as [Description]         
      ,CAST([IService1]             AS varchar) as [IService1]           
      ,CAST([IService2]             AS varchar) as [IService2]           
      ,CAST([LoopKEY]               AS varchar) as [LoopKEY]             
      ,CAST([OrderBy]               AS varchar) as [OrderBy]             
      ,CAST([LoopDwgCode]           AS varchar) as [LoopDwgCode]         
      ,CAST([ISALocation]           AS varchar) as [ISALocation]         
      ,CAST([ProjectStatus]         AS varchar) as [ProjectStatus]       
      ,CAST([PIDNo]                 AS varchar) as [PIDNo]               
      ,CAST([PlantConnection]       AS varchar) as [PlantConnection]     
      ,CAST([DCSIO]                 AS varchar) as [DCSIO]               
      ,CAST([DCSIOStatus]           AS varchar) as [DCSIOStatus]         
      ,CAST([DCSIOLocation]         AS varchar) as [DCSIOLocation]       
      ,CAST([PLCIO]                 AS varchar) as [PLCIO]               
      ,CAST([PLCIOStatus]           AS varchar) as [PLCIOStatus]         
      ,CAST([PLCIOLocation]         AS varchar) as [PLCIOLocation]       
      ,CAST([OtherIO]               AS varchar) as [OtherIO]             
      ,CAST([OperatingPrinc]        AS varchar) as [OperatingPrinc]      
      ,CAST([DataSheetNo]           AS varchar) as [DataSheetNo]         
      ,CAST([RequisitionNo]         AS varchar) as [RequisitionNo]       
      ,CAST([Manufacturer]          AS varchar) as [Manufacturer]        
      ,CAST([ModelNo]               AS varchar) as [ModelNo]             
      ,CAST([Supplier]              AS varchar) as [Supplier]            
      ,CAST([RangeLow]              AS varchar) as [RangeLow]            
      ,CAST([RangeHigh]             AS varchar) as [RangeHigh]           
      ,CAST([RangeUnits]            AS varchar) as [RangeUnits]          
      ,CAST([CalibratedRangeLow]    AS varchar) as [CalibratedRangeLow]  
      ,CAST([CalibratedRangeHigh]   AS varchar) as [CalibratedRangeHigh] 
      ,CAST([CalibratedRangeUnits]  AS varchar) as [CalibratedRangeUnits]
      ,CAST([SetPoint]              AS varchar) as [SetPoint]            
      ,CAST([IngressProtection]     AS varchar) as [IngressProtection]   
      ,CAST([Cost]                  AS varchar) as [Cost]                
      ,CAST([InstallCost]           AS varchar) as [InstallCost]         
      ,CAST([WiringConfig]          AS varchar) as [WiringConfig]        
      ,CAST([OldTagNo]              AS varchar) as [OldTagNo]            
      ,CAST([LoopDwgNo]             AS varchar) as [LoopDwgNo]           
      ,CAST([JunctionBox]           AS varchar) as [JunctionBox]         
      ,CAST([HookUpReqd]            AS varchar) as [HookUpReqd]          
      ,CAST([HookUpType]            AS varchar) as [HookUpType]          
      ,CAST([LocationDwg]           AS varchar) as [LocationDwg]         
      ,CAST([InstallDwgNo]          AS varchar) as [InstallDwgNo]        
      ,CAST([InstallationID]        AS varchar) as [InstallationID]      
      ,CAST([PlantKEY]              AS varchar) as [PlantKEY]            
      ,CAST([PlantObjectType]       AS varchar) as [PlantObjectType]     
      ,CAST([SignalType]            AS varchar) as [SignalType]          
      ,CAST([SignalVoltage]         AS varchar) as [SignalVoltage]       
      ,CAST([Size]                  AS varchar) as [Size]                
      ,CAST([ExRating]              AS varchar) as [ExRating]            
      ,CAST([PowerSupply]           AS varchar) as [PowerSupply]         
      ,CAST([Remarks]               AS varchar) as [Remarks]             
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList] 
  WHERE [TagNo] like 'ASBJ%')
  AS s UNPIVOT
  (
   [PropVal] FOR [PropName]
  IN(
    [ISAFunc]              
   ,[LNumber]              
   ,[Suffix]               
   ,[Prefix]               
   ,[Description]          
   ,[IService1]            
   ,[IService2]            
   ,[LoopKEY]              
   ,[OrderBy]              
   ,[LoopDwgCode]          
   ,[ISALocation]          
   ,[ProjectStatus]        
   ,[PIDNo]                
   ,[PlantConnection]      
   ,[DCSIO]                
   ,[DCSIOStatus]          
   ,[DCSIOLocation]        
   ,[PLCIO]                
   ,[PLCIOStatus]          
   ,[PLCIOLocation]        
   ,[OtherIO]              
   ,[OperatingPrinc]       
   ,[DataSheetNo]          
   ,[RequisitionNo]        
   ,[Manufacturer]         
   ,[ModelNo]              
   ,[Supplier]             
   ,[RangeLow]             
   ,[RangeHigh]            
   ,[RangeUnits]           
   ,[CalibratedRangeLow]   
   ,[CalibratedRangeHigh]  
   ,[CalibratedRangeUnits] 
   ,[SetPoint]             
   ,[IngressProtection]    
   ,[Cost]                 
   ,[InstallCost]          
   ,[WiringConfig]         
   ,[OldTagNo]             
   ,[LoopDwgNo]            
   ,[JunctionBox]          
   ,[HookUpReqd]           
   ,[HookUpType]           
   ,[LocationDwg]          
   ,[InstallDwgNo]         
   ,[InstallationID]       
   ,[PlantKEY]             
   ,[PlantObjectType]      
   ,[SignalType]           
   ,[SignalVoltage]        
   ,[Size]                 
   ,[ExRating]             
   ,[PowerSupply]          
   ,[Remarks]              
   )
) AS unpvt;"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC13_EI_Properties.csv" -NoTypeInformation
$finished = $true
Write-Log -Level INFO -Message "EL Props Extraction for $epc finished successfully." -finished $finished
  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
}

}


if ($epc -eq '11') {

  $finished = $false

  try {

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT [TagNo] as TagID
	,[PropName] as AttributeID
    ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData 
  ON [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList].[InstKEY] = [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData.[ObjectId]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef]
  ON [RYA-EI-TE-XXX-MCD].[dbo].IPropertyData.[PropertyDefID] = [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef].[PropertyDefID]
  WHERE [TagNo] NOT like 'ASBJ%'

UNION

SELECT [EquipmentNo] as TagID
	  ,[PropName] as AttributeID
      ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData] 
  ON [RYA-EI-TE-XXX-MCD].[dbo].[Equipment].[EquipId] = [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[ObjectId]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef]
  ON [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[PropertyDefID] = [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef].[PropertyDefID]
  WHERE [EquipmentNo] NOT like 'ASBJ%'

UNION
  
  
SELECT [TagNo] as TagID
      ,[PropName] as AttributeID
      ,[PropVal] as AttributeValue
	  ,'' as UnitOfMeasureID	
FROM (
SELECT      [TagNo]
      ,CAST([ISAFunc]               AS varchar) as [ISAFunc]             
      ,CAST([LNumber]               AS varchar) as [LNumber]             
      ,CAST([Suffix]                AS varchar) as [Suffix]              
      ,CAST([Prefix]                AS varchar) as [Prefix]              
      ,CAST([Description]           AS varchar) as [Description]         
      ,CAST([IService1]             AS varchar) as [IService1]           
      ,CAST([IService2]             AS varchar) as [IService2]           
      ,CAST([LoopKEY]               AS varchar) as [LoopKEY]             
      ,CAST([OrderBy]               AS varchar) as [OrderBy]             
      ,CAST([LoopDwgCode]           AS varchar) as [LoopDwgCode]         
      ,CAST([ISALocation]           AS varchar) as [ISALocation]         
      ,CAST([ProjectStatus]         AS varchar) as [ProjectStatus]       
      ,CAST([PIDNo]                 AS varchar) as [PIDNo]               
      ,CAST([PlantConnection]       AS varchar) as [PlantConnection]     
      ,CAST([DCSIO]                 AS varchar) as [DCSIO]               
      ,CAST([DCSIOStatus]           AS varchar) as [DCSIOStatus]         
      ,CAST([DCSIOLocation]         AS varchar) as [DCSIOLocation]       
      ,CAST([PLCIO]                 AS varchar) as [PLCIO]               
      ,CAST([PLCIOStatus]           AS varchar) as [PLCIOStatus]         
      ,CAST([PLCIOLocation]         AS varchar) as [PLCIOLocation]       
      ,CAST([OtherIO]               AS varchar) as [OtherIO]             
      ,CAST([OperatingPrinc]        AS varchar) as [OperatingPrinc]      
      ,CAST([DataSheetNo]           AS varchar) as [DataSheetNo]         
      ,CAST([RequisitionNo]         AS varchar) as [RequisitionNo]       
      ,CAST([Manufacturer]          AS varchar) as [Manufacturer]        
      ,CAST([ModelNo]               AS varchar) as [ModelNo]             
      ,CAST([Supplier]              AS varchar) as [Supplier]            
      ,CAST([RangeLow]              AS varchar) as [RangeLow]            
      ,CAST([RangeHigh]             AS varchar) as [RangeHigh]           
      ,CAST([RangeUnits]            AS varchar) as [RangeUnits]          
      ,CAST([CalibratedRangeLow]    AS varchar) as [CalibratedRangeLow]  
      ,CAST([CalibratedRangeHigh]   AS varchar) as [CalibratedRangeHigh] 
      ,CAST([CalibratedRangeUnits]  AS varchar) as [CalibratedRangeUnits]
      ,CAST([SetPoint]              AS varchar) as [SetPoint]            
      ,CAST([IngressProtection]     AS varchar) as [IngressProtection]   
      ,CAST([Cost]                  AS varchar) as [Cost]                
      ,CAST([InstallCost]           AS varchar) as [InstallCost]         
      ,CAST([WiringConfig]          AS varchar) as [WiringConfig]        
      ,CAST([OldTagNo]              AS varchar) as [OldTagNo]            
      ,CAST([LoopDwgNo]             AS varchar) as [LoopDwgNo]           
      ,CAST([JunctionBox]           AS varchar) as [JunctionBox]         
      ,CAST([HookUpReqd]            AS varchar) as [HookUpReqd]          
      ,CAST([HookUpType]            AS varchar) as [HookUpType]          
      ,CAST([LocationDwg]           AS varchar) as [LocationDwg]         
      ,CAST([InstallDwgNo]          AS varchar) as [InstallDwgNo]        
      ,CAST([InstallationID]        AS varchar) as [InstallationID]      
      ,CAST([PlantKEY]              AS varchar) as [PlantKEY]            
      ,CAST([PlantObjectType]       AS varchar) as [PlantObjectType]     
      ,CAST([SignalType]            AS varchar) as [SignalType]          
      ,CAST([SignalVoltage]         AS varchar) as [SignalVoltage]       
      ,CAST([Size]                  AS varchar) as [Size]                
      ,CAST([ExRating]              AS varchar) as [ExRating]            
      ,CAST([PowerSupply]           AS varchar) as [PowerSupply]         
      ,CAST([Remarks]               AS varchar) as [Remarks]             
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList] 
  WHERE [TagNo] NOT like 'ASBJ%')
  AS s UNPIVOT
  (
   [PropVal] FOR [PropName]
  IN(
    [ISAFunc]              
   ,[LNumber]              
   ,[Suffix]               
   ,[Prefix]               
   ,[Description]          
   ,[IService1]            
   ,[IService2]            
   ,[LoopKEY]              
   ,[OrderBy]              
   ,[LoopDwgCode]          
   ,[ISALocation]          
   ,[ProjectStatus]        
   ,[PIDNo]                
   ,[PlantConnection]      
   ,[DCSIO]                
   ,[DCSIOStatus]          
   ,[DCSIOLocation]        
   ,[PLCIO]                
   ,[PLCIOStatus]          
   ,[PLCIOLocation]        
   ,[OtherIO]              
   ,[OperatingPrinc]       
   ,[DataSheetNo]          
   ,[RequisitionNo]        
   ,[Manufacturer]         
   ,[ModelNo]              
   ,[Supplier]             
   ,[RangeLow]             
   ,[RangeHigh]            
   ,[RangeUnits]           
   ,[CalibratedRangeLow]   
   ,[CalibratedRangeHigh]  
   ,[CalibratedRangeUnits] 
   ,[SetPoint]             
   ,[IngressProtection]    
   ,[Cost]                 
   ,[InstallCost]          
   ,[WiringConfig]         
   ,[OldTagNo]             
   ,[LoopDwgNo]            
   ,[JunctionBox]          
   ,[HookUpReqd]           
   ,[HookUpType]           
   ,[LocationDwg]          
   ,[InstallDwgNo]         
   ,[InstallationID]       
   ,[PlantKEY]             
   ,[PlantObjectType]      
   ,[SignalType]           
   ,[SignalVoltage]        
   ,[Size]                 
   ,[ExRating]             
   ,[PowerSupply]          
   ,[Remarks]              
   )
) AS unpvt;"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC11_EI_Properties.csv" -NoTypeInformation

#GDP Not automatic upload for now
# $DataSet.Tables[0] | Export-Csv -Path "\\qamv3-sapp243\GDP\GDP_StagingArea\MP\AENG_PROPERTIES\EPCIC11_EI_Properties.csv" -NoTypeInformation

$finished = $true
Write-Log -Level INFO -Message "EL Props Extraction for $epc finished successfully." -finished $finished
  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
}

}

## EPC 12  Electrical ##

if ($epc -eq '12') {

  $finished = $false

  try {

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-BHA-LTM; Integrated Security = True;"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT [TagNo] as TagID
	,[PropName] as AttributeID
    ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-BHA-LTM].[dbo].[InstrumentList]
  LEFT JOIN [RYA-EI-TE-BHA-LTM].[dbo].IPropertyData 
  ON [RYA-EI-TE-BHA-LTM].[dbo].[InstrumentList].[InstKEY] = [RYA-EI-TE-BHA-LTM].[dbo].IPropertyData.[ObjectId]
  LEFT JOIN [RYA-EI-TE-BHA-LTM].[dbo].[PropertyDef]
  ON [RYA-EI-TE-BHA-LTM].[dbo].IPropertyData.[PropertyDefID] = [RYA-EI-TE-BHA-LTM].[dbo].[PropertyDef].[PropertyDefID]


UNION

SELECT [EquipmentNo] as TagID
	  ,[PropName] as AttributeID
      ,replace(replace(replace([PropValue], CHAR(13), ' '), CHAR(10), ' '), CHAR(13) + CHAR(10), ' ') as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-BHA-LTM].[dbo].[Equipment]
  LEFT JOIN [RYA-EI-TE-BHA-LTM].[dbo].[EPropertyData] 
  ON [RYA-EI-TE-BHA-LTM].[dbo].[Equipment].[EquipId] = [RYA-EI-TE-BHA-LTM].[dbo].[EPropertyData].[ObjectId]
  LEFT JOIN [RYA-EI-TE-BHA-LTM].[dbo].[PropertyDef]
  ON [RYA-EI-TE-BHA-LTM].[dbo].[EPropertyData].[PropertyDefID] = [RYA-EI-TE-BHA-LTM].[dbo].[PropertyDef].[PropertyDefID]
"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()


$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC12_EI_Properties.csv" -NoTypeInformation

#AIM PUBLISH. For now not pushing Automatically.
# $DataSet.Tables[0] | Export-Csv -Path "\\qamv3-sapp243\GDP\GDP_StagingArea\MP\AENG_PROPERTIES\EPCIC12_EI_Properties.csv" -NoTypeInformation
$finished = $true
Write-Log -Level INFO -Message "EL Props Extraction for $epc finished successfully." -finished $finished
}

catch {
    Write-Log -Level ERROR -Message "Failed to export CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
}

}
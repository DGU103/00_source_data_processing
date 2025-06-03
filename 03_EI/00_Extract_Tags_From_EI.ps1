
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)


$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "EI"

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

if ($epc -eq '11') {

  $finished = $false

  try {

## EPC 11  Electrical ##
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT  [EquipmentNo] as [TagNo], 
						--[Description] as [Equipment / Instrument Type Description], 
						--[Service] as [Description], 
						--'' as [Equipment / Instrument Type Description], 
						--'' as [Description], 
						'ElectricalEquipment' as [BaseType],
						GETDATE() as [DATE]
						  FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
						  
						  where (EquipTypeID in (7,74,20,28,22,29,34,25,23,44,31,30,43) 
						  and EquipmentNo not like 'ASBJ%')
						  or (EquipTypeID in (8,2) 
							and EquipmentNo not like 'ASBJ%'
							and (EquipmentNo like '%-CP-%' 
								or EquipmentNo like '%-JBP-%' 
								or EquipmentNo like '%-JBT-%'
								or EquipmentNo like '%-JBL-%'))"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC11_AE&I_Tagged_Electrical.csv" -NoTypeInformation


Write-Log -Level INFO -Message "Electical Tag Extraction for $epc finished successfully."

  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export Electrical Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
}

try {

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

## EPC 11  Instrumentation ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "select distinct [TagNo], 
						              'Instruments' as [BaseType],
						          GETDATE() as [DATE]
    from(SELECT  [EquipmentNo] as [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentEquipment' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
    where EquipTypeID in (4,9,17) and EquipmentNo not like 'ASBJ%'
    or 
    (EquipTypeID in (8,2) and EquipmentNo not like 'ASBJ%' 
    and 
    (EquipmentNo like '%-CPJ-%' 
    or EquipmentNo like '%-JBE-%' 
    or EquipmentNo like '%-JBJ-%'
    or EquipmentNo like '%-JBS-%' 
    or EquipmentNo like '%-JBC-%'))


    UNION 
    SELECT [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentList' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
    WHERE [TagNo] not like 'ASBJ%'
    AND  [TagNo] not like '%DEM[0-9]%'
    ) as dtl"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC11_AE&I_Tagged_Instruments.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Instrument Tag Extraction for $epc finished successfully."
}

catch {
  Write-Log -Level ERROR -Message "Failed to export Instrument Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

try {

  $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

## EPC 11  Cables ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT  CableNo as [TagNo],
--CONCAT(Remarks,'. ', UserField1,'. ', UserField2,'. ', UserField3,'.')  as [Description],
'Cables' as [BaseType],
GETDATE() as [DATE]
FROM [RYA-EI-TE-XXX-MCD].[dbo].[Cables]
where  CableNo not like 'ASBJ%'"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC11_AE&I_Tagged_Cables.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Cables Tag Extraction for $epc finished successfully."
}

catch {
  Write-Log -Level ERROR -Message "Failed to export Cable Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

try {

  $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

## EPC 11  Loops ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "WITH InstrumentationCTE as (
select distinct [TagNo], 
						              'Instruments' as [BaseType],
						          GETDATE() as [DATE]
    from(SELECT  [EquipmentNo] as [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentEquipment' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
    where EquipTypeID in (4,9,17) and EquipmentNo not like 'ASBJ%'
    or 
    (EquipTypeID in (8,2) and EquipmentNo not like 'ASBJ%' 
    and 
    (EquipmentNo like '%-CPJ-%' 
    or EquipmentNo like '%-JBE-%' 
    or EquipmentNo like '%-JBJ-%'
    or EquipmentNo like '%-JBS-%' 
    or EquipmentNo like '%-JBC-%'))


    UNION 
    SELECT [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentList' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
    WHERE [TagNo] not like 'ASBJ%'
    AND  [TagNo] not like '%DEM[0-9]%'
    ) as dtl )
  SELECT L.[LoopNo] FROM [RYA-EI-TE-XXX-MCD].[dbo].[LoopList] as L
  LEFT JOIN InstrumentationCTE As I
  ON L.LoopNo COLLATE
  SQL_Latin1_General_CP1_CI_AS = I.TagNo
  Where I.TagNo IS NULL;"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC11_Loops.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Loops Tag Extraction for $epc finished successfully."
}

catch {
  Write-Log -Level ERROR -Message "Failed to export Loop Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

$finished = $true
Write-Log -Level INFO -Message "EI Tags Export for $epc is finished" -finished $finished

}

if ($epc -eq '13') {

  $finished = $false

  try {

    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

## EPC 13  Electrical ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT distinct [EquipmentNo] as [TagNo], 
						          GETDATE() as [DATE]
                      FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
                      where EquipTypeID in (7,74,20,28,22,29,34,25,23,44,31,30,43) 
                      and EquipmentNo like 'ASBJ%'
                      or (EquipTypeID in (8,2) 
	                    and EquipmentNo like 'ASBJ%'
	                    and (EquipmentNo like '%-CP-%' 
		                    or EquipmentNo like '%-JBP-%' 
		                    or EquipmentNo like '%-JBT-%'
		                    or EquipmentNo like '%-JBL-%'))"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC13_AE&I_Tagged_Electrical.csv" -NoTypeInformation
Write-Log -Level INFO -Message "Electrical Tag Extraction for $epc finished successfully."
  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export Electrical Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
  }

  try {

    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"

## EPC 13  Instrumentation ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "select distinct [TagNo], 
    [BaseType],
		GETDATE() as [DATE]
    from(SELECT  [EquipmentNo] as [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentEquipment' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
    where EquipTypeID in (4,9,17) and EquipmentNo like 'ASBJ%'
    or 
    (EquipTypeID in (8,2) and EquipmentNo like 'ASBJ%' 
    and 
    (EquipmentNo like '%-CPJ-%' 
    or EquipmentNo like '%-JBE-%' 
    or EquipmentNo like '%-JBJ-%'
    or EquipmentNo like '%-JBS-%' 
    or EquipmentNo like '%-JBC-%'))

    UNION 
    SELECT [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentList' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
    WHERE [TagNo] like 'ASBJ%'
    AND  [TagNo] not like '%DEM[0-9]%'
    ) as dtl"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC13_AE&I_Tagged_Instruments.csv" -NoTypeInformation
Write-Log -Level INFO -Message "Instrumentaiton Tag Extraction for $epc finished successfully."
  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export Instrumentaiton Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
  }

  try {
  
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"
## EPC 13  Cables ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT  CableNo as [TagNo],
--CONCAT(Remarks,'. ', UserField1,'. ', UserField2,'. ', UserField3,'.')  as [Description],
'Cables' as [BaseType],
GETDATE() as [DATE]
FROM [RYA-EI-TE-XXX-MCD].[dbo].[Cables]
where  CableNo like 'ASBJ%'"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC13_AE&I_Tagged_Cables.csv" -NoTypeInformation
Write-Log -Level INFO -Message "Cable Tag Extraction for $epc finished successfully."

  }

  catch {
    Write-Log -Level ERROR -Message "Failed to export Cable Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
    throw
  }

  try {

## EPC 13  Loops ##
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-XXX-MCD; Integrated Security = True;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "WITH InstrumentationCTE as (
select distinct [TagNo], 
    [BaseType],
		GETDATE() as [DATE]
    from(SELECT  [EquipmentNo] as [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentEquipment' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
    where EquipTypeID in (4,9,17) and EquipmentNo like 'ASBJ%'
    or 
    (EquipTypeID in (8,2) and EquipmentNo like 'ASBJ%' 
    and 
    (EquipmentNo like '%-CPJ-%' 
    or EquipmentNo like '%-JBE-%' 
    or EquipmentNo like '%-JBJ-%'
    or EquipmentNo like '%-JBS-%' 
    or EquipmentNo like '%-JBC-%'))

    UNION 
    SELECT [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentList' as [BaseType]

    FROM [RYA-EI-TE-XXX-MCD].[dbo].[InstrumentList]
    WHERE [TagNo] like 'ASBJ%'
    AND  [TagNo] not like '%DEM[0-9]%'
    ) as dtl )

  SELECT L.[LoopNo] FROM [RYA-EI-TE-XXX-MCD].[dbo].[LoopList] as L
  LEFT JOIN InstrumentationCTE As I
  ON L.LoopNo COLLATE
  SQL_Latin1_General_CP1_CI_AS = I.TagNo
  Where I.TagNo IS NULL;"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC13_Loops.csv" -NoTypeInformation
Write-Log -Level INFO -Message "Loops Tag Extraction for $epc finished successfully."

}

catch {
  Write-Log -Level ERROR -Message "Failed to export Cable Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

$finished = $true
Write-Log -Level INFO -Message "EI Tags Export for $epc is finished" -finished $finished

}

if ($epc -eq '12') {

  $finished = $false

  try {

## EPCI 12 Electrical ##
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-BHA-LTM; Integrated Security = True;"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT distinct [EquipmentNo] as [TagNo], 
						            GETDATE() as [DATE]
                        FROM [RYA-EI-TE-BHA-LTM].[dbo].[Equipment]
                        where EquipTypeID in (7,74,20,28,22,29,34,25,23,44,31,30,43) 
                        or (EquipTypeID in (8,2) 
		                        and (EquipmentNo like '%-CP-%' 
		                        or EquipmentNo like '%-JBP-%' 
		                        or EquipmentNo like '%-JBT-%'
		                        or EquipmentNo like '%-JBL-%'))"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC12_AE&I_Tagged_Electrical.csv" -NoTypeInformation
Write-Log -Level INFO -Message "Electrical Tag Extraction for $epc finished successfully."

}

catch {
  Write-Log -Level ERROR -Message "Failed to export Electrical Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

try {

  $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-BHA-LTM; Integrated Security = True;"
## EPCI 12 Instrumentation ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = " select distinct [TagNo], 
                         --[Equipment / Instrument Type Description], 
                         --[Description],
						              'Instruments' as [BaseType],
						              GETDATE() as [DATE]
                         from(
                         SELECT [EquipmentNo] as [TagNo], 
                         '' as [Equipment / Instrument Type Description], 
                         '' as [Description],
						 'InstrumentEquipment' as [BaseType]
                         FROM [RYA-EI-TE-BHA-LTM].[dbo].[Equipment]
                         where EquipTypeID in (4,9,17)
                         or (EquipTypeID in (8,2) 
                         and (EquipmentNo like '%-CPJ-%' 
                         or EquipmentNo like '%-JBE-%' 
                         or EquipmentNo like '%-JBJ-%'
                         or EquipmentNo like '%-JBS-%' 
                         or EquipmentNo like '%-JBC-%'))
                         UNION 
                         select [TagNo], 
                         '' as [Equipment / Instrument Type Description], 
                         '' as [Description],
						 'InstrumentList' as [BaseType]
                         from [RYA-EI-TE-BHA-LTM].[dbo].InstrumentList) as dtl"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC12_AE&I_Tagged_Instruments.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Instrumentation Tag Extraction for $epc finished successfully."

}

catch {
  Write-Log -Level ERROR -Message "Failed to export Instrumentation Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

try {
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-BHA-LTM; Integrated Security = True;"
## EPC 12  Cables ##
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "SELECT  CableNo as [TagNo],
--CONCAT(Remarks,'. ', UserField1,'. ', UserField2,'. ', UserField3,'.')  as [Description],
'Cables' as [BaseType],
GETDATE() as [DATE]
FROM [RYA-EI-TE-BHA-LTM].[dbo].[Cables]"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC12_AE&I_Tagged_Cables.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Cables Tag Extraction for $epc finished successfully."

}

catch {
  Write-Log -Level ERROR -Message "Failed to export Cables Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

try {
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = QA-SQL-TEST2019; Database = RYA-EI-TE-BHA-LTM; Integrated Security = True;"
## EPC 12  Loops ##

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "WITH InstrumentationCTE as (
select distinct [TagNo], 'Instruments' as [BaseType],
						          GETDATE() as [DATE]
    from(SELECT  [EquipmentNo] as [TagNo], 
    '' as [Equipment / Instrument Type Description], 
    '' as [Description],
    'InstrumentEquipment' COLLATE SQL_Latin1_General_CP1_CI_AS  as [BaseType]

                  FROM [RYA-EI-TE-BHA-LTM].[dbo].[Equipment]
                        where EquipTypeID in (4,9,17)
                         or (EquipTypeID in (8,2) 
                         and (EquipmentNo like '%-CPJ-%' 
                         or EquipmentNo like '%-JBE-%' 
                         or EquipmentNo like '%-JBJ-%'
                         or EquipmentNo like '%-JBS-%' 
                         or EquipmentNo like '%-JBC-%'))
                         UNION 
                         select [TagNo], 
                         '' as [Equipment / Instrument Type Description], 
                         '' as [Description],
						 'InstrumentList' as [BaseType]
                         from [RYA-EI-TE-BHA-LTM].[dbo].InstrumentList) as dtl )
  
  SELECT L.[LoopNo] FROM [RYA-EI-TE-BHA-LTM].[dbo].[LoopList] as L
  LEFT JOIN InstrumentationCTE As I
  ON L.LoopNo COLLATE
  SQL_Latin1_General_CP1_CI_AS = I.TagNo
  Where I.TagNo IS NULL;
"

$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()

$DataSet.Tables[0] | Export-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\EI\EPCIC12_Loops.csv" -NoTypeInformation

Write-Log -Level INFO -Message "Loops Tag Extraction for $epc finished successfully."

}

catch {
  Write-Log -Level ERROR -Message "Failed to export Loops Tag CSV for EPCIC $epc. Error: $($_.Exception.Message)"
  throw
}

$finished = $true
Write-Log -Level INFO -Message "EI Tags Export for $epc is finished" -finished $finished

}
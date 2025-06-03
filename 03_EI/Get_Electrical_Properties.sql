/****** Script for SelectTopNRows command from SSMS  ******/
SELECT [EquipmentNo] as TagID
	  ,[PropName] as AttributeID
      ,[PropValue] as AttributeValue
	,'' as UnitOfMeasureID	
  FROM [RYA-EI-TE-XXX-MCD].[dbo].[Equipment]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData] 
  ON [RYA-EI-TE-XXX-MCD].[dbo].[Equipment].[EquipId] = [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[ObjectId]
  LEFT JOIN [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef]
  ON [RYA-EI-TE-XXX-MCD].[dbo].[EPropertyData].[PropertyDefID] = [RYA-EI-TE-XXX-MCD].[dbo].[PropertyDef].[PropertyDefID]




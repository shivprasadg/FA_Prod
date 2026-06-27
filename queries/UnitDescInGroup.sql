SELECT
  LookUp_UnitDesc.GroupID,
  Count(LookUp_UnitDesc.AssetID) AS Units,
  "(" & Count([assetid])& ") " & [DescriptionShort] AS UnitDesc,
  LookUp_UnitDesc.DescriptionShort
FROM
  LookUp_UnitDesc
GROUP BY
  LookUp_UnitDesc.GroupID,
  LookUp_UnitDesc.DescriptionShort;

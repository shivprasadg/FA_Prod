SELECT
  LookUp_UnitDesc.GroupID,
  Count(LookUp_UnitDesc.AssetID) AS Units,
  "(" & Count([assetid])& ") " & IIf(
    Len([LegalDescription])< 3,
    [DescriptionLong],
    [LegalDescription]
  ) AS UnitDesc,
  LookUp_UnitDesc.DescriptionLong,
  LookUp_UnitDesc.LegalDescription
FROM
  LookUp_UnitDesc
GROUP BY
  LookUp_UnitDesc.GroupID,
  LookUp_UnitDesc.DescriptionLong,
  LookUp_UnitDesc.LegalDescription;

SELECT
  LookUp_UnitDesc.GroupID,
  LookUp_UnitDesc.DescAbbrv AS ShortDesc,
  LookUp_UnitDesc.DescriptionShort
FROM
  LookUp_UnitDesc
GROUP BY
  LookUp_UnitDesc.GroupID,
  LookUp_UnitDesc.DescAbbrv,
  LookUp_UnitDesc.DescriptionShort;

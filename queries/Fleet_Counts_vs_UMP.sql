SELECT
  Min(vw_SixKeys.ClientGroupId) AS ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Status,
  Count(vw_SixKeys.Unitid) AS Units,
  Count(UMP_DataCount.UnitId) AS HasOBCData,
  [Units] - [HasOBCData] AS NoOBCData,
  vw_SixKeys.CompanyType
FROM
  vw_SixKeys
  LEFT JOIN UMP_DataCount ON vw_SixKeys.UnitID = UMP_DataCount.UnitId
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Status,
  vw_SixKeys.CompanyType;

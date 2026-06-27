SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.Unitnum,
  vw_SixKeys.Status,
  vw_SixKeys.MYear,
  [Make] & " - " & [Model] & " - " & [type] AS [Desc],
  Units.UnitCOABack,
  Switch(
    IsDate([UnitCOABack])
    And [Status] = "B",
    "Should be 'A'"
  ) AS Error
FROM
  vw_SixKeys
  INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
WHERE
  (
    (
      (vw_SixKeys.Status) Not Like "*K*"
      And (vw_SixKeys.Status) Not Like "*R*"
    )
    AND (
      (Units.UnitCOABack) Is Not Null
    )
  )
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup;

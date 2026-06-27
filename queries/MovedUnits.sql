SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.SchID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum,
  vw_SixKeys.VIN,
  UnitLocationChangeHistory.LocationChangeDate,
  UnitLocationChangeHistory.PrevLocation,
  UnitLocationChangeHistory.NewLocation,
  UnitLocationChangeHistory.Reason
FROM
  vw_SixKeys
  INNER JOIN UnitLocationChangeHistory ON vw_SixKeys.AssetID = UnitLocationChangeHistory.AssetId
WHERE
  (
    (
      (vw_SixKeys.SchID)= 846
    )
  )
ORDER BY
  UnitLocationChangeHistory.LocationChangeDate;

SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum,
  vw_SixKeys.UnitDescShortVer,
  Units.ObcUnitNumber,
  vw_SixKeys.AcceptedDate,
  vw_SixKeys.InServiceDate
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (Units.ObcUnitNumber) Is Null
    )
    AND (
      (Units.UnitStatus)<> "B"
    )
  );

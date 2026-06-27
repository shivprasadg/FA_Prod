SELECT
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Units.GroupREF,
  Units.UnitID,
  Units.UnitVIN,
  Units.unitstatus,
  vw_SixKeys.UnitGroup
FROM
  vw_SixKeys
  INNER JOIN (
    Units
    LEFT JOIN UnitVinDecoder ON Units.UnitVIN = UnitVinDecoder.VIN
  ) ON vw_SixKeys.AssetID = Units.UnitID
WHERE
  (
    (
      (Units.unitstatus)= "A"
    )
    AND (
      (UnitVinDecoder.VIN) Is Null
    )
  );

SELECT
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Units.GroupREF,
  vw_SixKeys.VIN,
  Units.UnitID,
  Units.unitstatus,
  UnitVinDecoder.*,
  UnitVinDecoder.VEHICLE_TYPE AS Expr1
FROM
  vw_SixKeys
  INNER JOIN (
    Units
    LEFT JOIN UnitVinDecoder ON Units.UnitVIN = UnitVinDecoder.VIN
  ) ON vw_SixKeys.AssetID = Units.UnitID
WHERE
  (
    (
      (Units.UnitID)= 275946
      Or (Units.UnitID)= 276013
      Or (Units.UnitID)= 276014
      Or (Units.UnitID)= 276015
      Or (Units.UnitID)= 276016
      Or (Units.UnitID)= 286933
      Or (Units.UnitID)= 286932
      Or (Units.UnitID)= 286931
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;

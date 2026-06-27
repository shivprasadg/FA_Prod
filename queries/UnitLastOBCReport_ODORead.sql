SELECT
  UnitLastOBCReport.UnitID,
  vw_UMP_Mileage.LogActivityDate AS LastRead,
  vw_UMP_Mileage.OdometerBegin,
  vw_UMP_Mileage.OdometerEnd,
  vw_UMP_Mileage.ClientID
FROM
  UnitLastOBCReport
  INNER JOIN vw_UMP_Mileage ON (
    UnitLastOBCReport.LogActivityDate = vw_UMP_Mileage.LogActivityDate
  )
  AND (
    UnitLastOBCReport.UnitID = vw_UMP_Mileage.UnitID
  );

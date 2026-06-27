SELECT
  UnitFirstOBCReport.UnitID,
  vw_UMP_Mileage.LogActivityDate AS FirstRead,
  vw_UMP_Mileage.OdometerBegin,
  vw_UMP_Mileage.OdometerEnd,
  vw_UMP_Mileage.ClientID
FROM
  UnitFirstOBCReport
  INNER JOIN vw_UMP_Mileage ON (
    UnitFirstOBCReport.UnitID = vw_UMP_Mileage.UnitID
  )
  AND (
    UnitFirstOBCReport.LogActivityDate = vw_UMP_Mileage.LogActivityDate
  );

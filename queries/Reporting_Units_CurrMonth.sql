SELECT
  vw_UMP_Mileage.UnitID,
  vw_UMP_Mileage.UnitNumber,
  vw_UMP_Mileage.OBCUnitNumber,
  vw_UMP_Mileage.LogActivityDate,
  vw_UMP_Mileage.ReportYear,
  vw_UMP_Mileage.ReportMonth,
  vw_UMP_Mileage.ClientGroupID,
  vw_UMP_Mileage.ClientID,
  vw_SixKeys.VIN,
  vw_SixKeys.Status,
  vw_SixKeys.SubType
FROM
  vw_UMP_Mileage
  INNER JOIN vw_SixKeys ON vw_UMP_Mileage.UnitID = vw_SixKeys.AssetID
WHERE
  (
    (
      (vw_UMP_Mileage.ReportYear)= 2016
    )
    AND (
      (vw_UMP_Mileage.ReportMonth)= 8
    )
    AND (
      (vw_UMP_Mileage.ClientGroupID)= 60
    )
    AND (
      (vw_SixKeys.Status)= "A"
      Or (vw_SixKeys.Status)= "T"
    )
    AND (
      (vw_SixKeys.SubType)= "DayCab"
      Or (vw_SixKeys.SubType)= "Sleeper"
    )
  );

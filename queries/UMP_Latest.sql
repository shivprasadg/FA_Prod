SELECT
  Max(vw_UMP_Mileage.LogActivityDate) AS LogActivityDate,
  vw_UMP_Mileage.ClientGroupId,
  vw_UMP_Mileage.ClientId,
  vw_UMP_Mileage.UnitId,
  vw_UMP_Mileage.UnitNumber AS Expr1,
  Min(vw_UMP_Mileage.OBCDataStatus) AS OBCDataStatus
FROM
  vw_UMP_Mileage
GROUP BY
  vw_UMP_Mileage.ClientGroupId,
  vw_UMP_Mileage.ClientId,
  vw_UMP_Mileage.UnitId,
  vw_UMP_Mileage.UnitNumber
HAVING
  (
    (
      (
        Max(vw_UMP_Mileage.LogActivityDate)
      )= #7/1/2015#
    )
    AND (
      (vw_UMP_Mileage.ClientGroupId)= 60
    )
  );

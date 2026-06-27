SELECT
  vw_UMP_Mileage.UnitID,
  Max(vw_UMP_Mileage.LogActivityDate) AS LogActivityDate,
  vw_UMP_Mileage.ClientID,
  vw_UMP_Mileage.ClientGroupID
FROM
  vw_UMP_Mileage
WHERE
  (
    (
      (vw_UMP_Mileage.OBCDataStatus) Not Like "*1792*"
    )
  )
GROUP BY
  vw_UMP_Mileage.UnitID,
  vw_UMP_Mileage.ClientID,
  vw_UMP_Mileage.ClientGroupID
HAVING
  (
    (
      (vw_UMP_Mileage.ClientGroupID)<> 435
    )
  );

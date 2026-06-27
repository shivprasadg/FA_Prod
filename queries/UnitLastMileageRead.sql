SELECT
  Units.UnitID,
  Units.UnitYr AS ModelYear,
  Max(Mileage.LogActivityDate) AS ReadDate,
  Max(
    Nz([OdometerEnd], 0)
  ) AS Mileage,
  Max(
    Nz([HoursLTD], 0)
  ) AS Hours
FROM
  Mileage
  RIGHT JOIN Units ON Mileage.UnitId = Units.UnitID
WHERE
  (
    (
      (Mileage.OBCDataStatus)= 1
      Or (Mileage.OBCDataStatus)= 0
    )
  )
GROUP BY
  Units.UnitID,
  Units.UnitYr;

SELECT
  Mileage.UnitID,
  Max(Mileage.MHDate) AS MHDate,
  Max(Mileage.AquiMI) AS Odometer
FROM
  Mileage
WHERE
  (
    (
      (Mileage.UnitID) Is Not Null
    )
  )
GROUP BY
  Mileage.UnitID
ORDER BY
  Mileage.UnitID;

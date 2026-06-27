SELECT
  Mileage.UnitID,
  Max(Mileage.MHDate) AS MaxOfMHDate
FROM
  Mileage
GROUP BY
  Mileage.UnitID
ORDER BY
  Mileage.UnitID;

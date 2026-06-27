SELECT
  Mileage.UnitID,
  Mileage.MHDate AS Expr1,
  Mileage.AquiMI AS Expr2,
  Mileage.AqiHr AS Expr3,
  Mileage.Mileage AS Expr4,
  Mileage.Hours
FROM
  Mileage
WHERE
  (
    (
      (Mileage.AquiMI)= Yes
    )
  )
  Or (
    (
      (Mileage.AqiHr)= Yes
    )
  )
ORDER BY
  Mileage.UnitID;

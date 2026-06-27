SELECT
  Count(Units.UnitVIN) AS CountOfUnitVIN,
  Units.UnitDocAccpt
FROM
  Units
WHERE
  (
    (
      (Units.UnitCOASent) Is Not Null
      And (Units.UnitCOASent)>= [enter start date]
      And (Units.UnitCOASent)<= [enter end date]
    )
  )
GROUP BY
  Units.UnitDocAccpt
HAVING
  (
    (
      (Units.UnitDocAccpt)= Yes
    )
  );

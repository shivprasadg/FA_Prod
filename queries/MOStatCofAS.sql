SELECT
  Count(Units.UnitVIN) AS CountOfUnitVIN
FROM
  Units
WHERE
  (
    (
      (Units.UnitCOASent) Is Not Null
      And (Units.UnitCOASent)>= [enter start date]
      And (Units.UnitCOASent)<= [enter end date]
    )
  );

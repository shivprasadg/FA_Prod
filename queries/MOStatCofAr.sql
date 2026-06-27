SELECT
  Count(Units.UnitVIN) AS CountOfUnitVIN
FROM
  Units
WHERE
  (
    (
      (Units.UnitCOABack) Is Not Null
      And (Units.UnitCOABack)>= [enter start date]
      And (Units.UnitCOABack)<= [enter end date]
    )
  );

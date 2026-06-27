SELECT
  Units.UnitID,
  Units.ClientGroupId,
  Units.UnitUnitNum,
  Units.ObcUnitNumber
FROM
  Units
WHERE
  (
    (
      (Units.ObcUnitNumber) Is Null
    )
  );

SELECT
  DISTINCT Units.VendorREF
FROM
  Units
WHERE
  (
    (
      (Units.UnitStatus)= "b"
    )
  )
  OR (
    (
      (Units.UnitStatus)= "a"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
      And (Units.UnitAcceptDate)> #12/30/2010#
    )
  )
  OR (
    (
      (Units.UnitStatus)= "a"
    )
    AND (
      (Units.UnitInServ) Is Not Null
      And (Units.UnitInServ)> #12/30/2010#
    )
  );

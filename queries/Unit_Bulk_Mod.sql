UPDATE
  DISTINCTROW (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
SET
  Units.UnitModel = "CA125",
  Units.VendorREF = 226,
  Units.UnitFinalCost = 113340,
  Units.UNITRENT = 1358.21,
  Units.UNITPDRENT = 45.27,
  Units.UnitSalesResid = 32000
WHERE
  (
    (
      (Schedule.SchFAID)= "14-652"
    )
    AND (
      (SchGrp.SGrpNO)= "A"
    )
    AND (
      (Schedule.SchNo)= "2"
    )
  );

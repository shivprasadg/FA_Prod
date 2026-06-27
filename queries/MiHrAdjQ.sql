SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  Units.UnitAcceptDate,
  Units.unitinserv,
  Units.unitunitnum,
  Mileage.MHDate AS Expr1,
  Mileage.Mileage AS Expr2,
  Mileage.Hours
FROM
  MstrLease
  INNER JOIN (
    Schedule
    INNER JOIN (
      SchGrp
      INNER JOIN (
        Mileage
        INNER JOIN Units ON Mileage.UnitID = Units.UnitID
      ) ON SchGrp.SGrpID = Units.GroupREF
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  ) ON MstrLease.MLID = Schedule.SchMLRef
WHERE
  (
    (
      (MstrLease.MLNo) Like "costco*"
    )
    And (
      (Schedule.SchNo)= "1 mira loma"
    )
    And (
      (Mileage.MHDate)>= "#8/1/2012#"
      And (Mileage.MHDate)<= "#9/1/2012#"
    )
    And (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Schedule.SchID,
  Units.unitunitnum,
  Mileage.MHDate;

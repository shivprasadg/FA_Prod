SELECT
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  IIf(
    IsNull([BLCD_Projected]),
    DateAdd("m", 1, [projIns])+(
      25 - DatePart(
        "d",
        DateAdd("m", 1, [projIns])
      )
    ),
    [BLCD_Projected]
  ) AS ProjBLCD
FROM
  (
    (
      (
        MstrLease
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN MaxDelTRFq ON SchGrp.SGrpID = MaxDelTRFq.SGrpID
GROUP BY
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  IIf(
    IsNull([BLCD_Projected]),
    DateAdd("m", 1, [projIns])+(
      25 - DatePart(
        "d",
        DateAdd("m", 1, [projIns])
      )
    ),
    [BLCD_Projected]
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo;

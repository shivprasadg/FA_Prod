SELECT
  Units.UnitTitleRecd,
  Units.UnitTASent,
  Schedule.SchTitleToTrust,
  Units.unitvin,
  SchGrp.Assignee_Bank,
  Units.UnitAcceptDate,
  CLASS.TITLED,
  MstrLease.mlorig,
  Units.UnitBuildDate,
  Units.UnitPOA,
  Units.UnitAddDt,
  SchGrp.sgrpplb
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
  LEFT JOIN CLASS ON Units.UnitType = CLASS.CLASS
WHERE
  (
    (
      (Units.UnitTitleRecd)>= [enter start dt]
      And (Units.UnitTitleRecd)<= [enter end dt]
    )
    AND (
      (CLASS.TITLED)= Yes
    )
    AND (
      (MstrLease.mlorig)= "fa"
    )
  )
ORDER BY
  Units.UnitTitleRecd;

SELECT
  SchGrp.UnitsInGroup,
  MstrLease.MLOrig,
  Schedule.SchNo,
  Units.unitstatus,
  Count(Units.unitvin) AS CountOfunitvin,
  Units.unitcoasent,
  IIf(
    IsNull([UnitBuildDate])
    Or [UnitBuildDate] > #12/31/2011#,
    "Yes",
    "No"
  ) AS 2012
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  SchGrp.UnitsInGroup,
  MstrLease.MLOrig,
  Schedule.SchNo,
  Units.unitstatus,
  Units.unitcoasent,
  IIf(
    IsNull([UnitBuildDate])
    Or [UnitBuildDate] > #12/31/2011#,
    "Yes",
    "No"
  ),
  MstrLease.MLNo
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (MstrLease.MLNo) Like "Maines*"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo;

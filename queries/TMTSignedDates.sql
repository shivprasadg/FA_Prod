SELECT
  Schedule.SchID,
  GroupFinalizedTracker.GroupID,
  SchGrp.SGrpNO AS GrpNo,
  [SGrpNO] & " - " & Format([RequestCompleted], "mm/dd/yy") AS TMTSigned
FROM
  Schedule
  LEFT JOIN (
    SchGrp
    LEFT JOIN GroupFinalizedTracker ON SchGrp.SGrpID = GroupFinalizedTracker.GroupID
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (
        GroupFinalizedTracker.Department
      )= "TMT"
      Or (
        GroupFinalizedTracker.Department
      ) Is Null
    )
  )
GROUP BY
  Schedule.SchID,
  GroupFinalizedTracker.GroupID,
  SchGrp.SGrpNO,
  [SGrpNO] & " - " & Format([RequestCompleted], "mm/dd/yy");

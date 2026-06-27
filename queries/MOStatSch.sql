SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchOut,
  Schedule.SchIn,
  SchGrp.AssignmentDate,
  SchGrp.Assignee_Bank,
  SchGrp.Assigned_Indicator,
  SchGrp.LeaseTermPrimary,
  SchGrp.FundDate,
  SchGrp.SaleDate
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (Schedule.SchIn) Is Not Null
      And (Schedule.SchIn)>= [enter start dt]
      And (Schedule.SchIn)<= [enter end date]
    )
  )
ORDER BY
  Schedule.SchOut;

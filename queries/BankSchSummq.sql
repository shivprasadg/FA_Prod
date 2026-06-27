SELECT
  MstrLease.MLOrig,
  MstrLease.MLID,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchFSL,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Schedule.SchOut,
  Schedule.SchIn,
  Schedule.SchAssDt,
  Schedule.SchShortDesc,
  Schedule.SchFAID,
  Sum(SchGrp.OrigEquipCost) AS SumOfOrigEquipCost,
  MstrLease.Client,
  Clients.ClientShNm,
  SchGrp.Assigned_Indicator,
  SchGrp.Assignee_Bank,
  SchGrp.BLCD,
  SchGrp.LeasePymtForGroup,
  SchGrp.LeaseType,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.AssignmentDate,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1 AS Exp,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  IIf(
    [ExchangeITGroup] = False, "none",
    "@ Mo " & [ExchangeITMonth]
  ) AS ExIT,
  SchGrp.Assignee_Bank,
  SchGrp.AssignmentDate
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
GROUP BY
  MstrLease.MLOrig,
  MstrLease.MLID,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchFSL,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Schedule.SchOut,
  Schedule.SchIn,
  Schedule.SchAssDt,
  Schedule.SchShortDesc,
  Schedule.SchFAID,
  MstrLease.Client,
  Clients.ClientShNm,
  SchGrp.Assigned_Indicator,
  SchGrp.BLCD,
  SchGrp.LeaseType,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  IIf(
    [ExchangeITGroup] = False, "none",
    "@ Mo " & [ExchangeITMonth]
  ),
  SchGrp.Assignee_Bank,
  SchGrp.AssignmentDate,
  SchGrp.LeasePymtForGroup,
  SchGrp.Assignee_Bank,
  SchGrp.AssignmentDate,
  SchGrp.LeasePymtForGroup,
  MstrLease.MLOrig,
  SchGrp.AssignmentDate
HAVING
  (
    (
      (MstrLease.MLNo) Not Like "*demo*"
    )
    AND (
      (SchGrp.Assignee_Bank) Is Not Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID DESC,
  Schedule.SchNo DESC,
  SchGrp.SGrpNO;

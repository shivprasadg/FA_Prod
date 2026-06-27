SELECT
  MstrLease.Client AS ClientID,
  Schedule.SchMLRef AS MLID,
  SchGrp.SGrpSchID AS ScheduleID,
  SchGrp.SGrpID AS GroupId,
  SchGrp.SGrpFAID AS FAID,
  SchGrp.BLCD AS BLCD,
  Min(SchGrp.GroupAddedDate) AS DateGroupAdded,
  Min(Schedule.DateSchAdded) AS DateSchAdded,
  SchGrp.SGrpNO AS [Group],
  Schedule.SchNo AS Schedule,
  Schedule.ApprovalRequest,
  Schedule.ApprovalReply
FROM
  MstrLease
  INNER JOIN (
    SchGrp
    INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
  ) ON MstrLease.MLID = Schedule.SchMLRef
GROUP BY
  MstrLease.Client,
  Schedule.SchMLRef,
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  SchGrp.SGrpFAID,
  SchGrp.BLCD,
  SchGrp.SGrpNO,
  Schedule.SchNo,
  Schedule.ApprovalRequest,
  Schedule.ApprovalReply;

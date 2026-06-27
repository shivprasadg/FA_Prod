SELECT
  Clients.clientID,
  MstrLease.MLOrig,
  MstrLease.MLID,
  MstrLease.MLNo,
  Schedule.SCHNo1,
  Schedule.SchID,
  Schedule.SchFSL,
  Schedule.SchMLRef,
  Schedule.SchNo,
  Schedule.SchAlias,
  Schedule.SchOut,
  Schedule.SchIn,
  Schedule.Schsigned,
  Schedule.SchSigner,
  Schedule.SchSigner2,
  Schedule.SchSigner3,
  Schedule.SchAssDt,
  Schedule.SchTitleAssignee,
  Schedule.SchTitleToTrust,
  Schedule.SchShortDesc,
  Schedule.SchFAID,
  Schedule.SCHINSAMT,
  Schedule.SCHINREC,
  Schedule.SCHINSEXP,
  Schedule.SchTitltoFA,
  Schedule.SchTitletoClient,
  Schedule.LienHolder,
  Schedule.LienHolderNm,
  MstrLease.MLLessor,
  MstrLease.Client,
  Clients.ClientShNm,
  "Edit" AS Edit,
  Schedule.SchAH,
  IIf(
    IsNull([Assignee_Bank]),
    Null,
    [Assignee_Bank]
  ) AS Bassign,
  IIf(
    IsNull([SumOfOrigEquipCost]),
    Null,
    [SumOfOrigEquipCost]
  ) AS ECost,
  Max(Units.UnitEstDelDate) AS MaxOfUnitEstDelDate,
  Max(
    IIf(
      IsNull([SgrpPlb]),
      No,
      [SgrpPlb]
    )
  ) AS PLB
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        Schedule
        LEFT JOIN SchEquipCostq ON Schedule.SchID = SchEquipCostq.SchID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    LEFT JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  Clients.clientID,
  MstrLease.MLOrig,
  MstrLease.MLID,
  MstrLease.MLNo,
  Schedule.SCHNo1,
  Schedule.SchID,
  Schedule.SchFSL,
  Schedule.SchMLRef,
  Schedule.SchNo,
  Schedule.SchAlias,
  Schedule.SchOut,
  Schedule.SchIn,
  Schedule.Schsigned,
  Schedule.SchSigner,
  Schedule.SchSigner2,
  Schedule.SchSigner3,
  Schedule.SchAssDt,
  Schedule.SchTitleAssignee,
  Schedule.SchTitleToTrust,
  Schedule.SchShortDesc,
  Schedule.SchFAID,
  Schedule.SCHINSAMT,
  Schedule.SCHINREC,
  Schedule.SCHINSEXP,
  Schedule.SchTitltoFA,
  Schedule.SchTitletoClient,
  Schedule.LienHolder,
  Schedule.LienHolderNm,
  MstrLease.MLLessor,
  MstrLease.Client,
  Clients.ClientShNm,
  "Edit",
  Schedule.SchAH,
  IIf(
    IsNull([Assignee_Bank]),
    Null,
    [Assignee_Bank]
  ),
  IIf(
    IsNull([SumOfOrigEquipCost]),
    Null,
    [SumOfOrigEquipCost]
  ),
  Schedule.SchFAID
HAVING
  (
    (
      (Clients.clientID)= [Forms]![ClientsView]![clientID]
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SCHNo1 DESC,
  Schedule.SchID DESC,
  Schedule.SchNo DESC;

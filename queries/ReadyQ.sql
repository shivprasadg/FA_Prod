SELECT
  DatePart("yyyy", [FundDate_Projected]) AS SynYr,
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLBack,
  MstrLease.MLSent,
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.UnitsInGroup,
  Schedule.SchShortDesc,
  Schedule.SchIn,
  Schedule.SchOut,
  SchGrp.SGrpNO,
  SchGrp.OrigEquipCost,
  SchGrp.Assignee_Bank,
  SchGrp.LeaseType,
  Sum(
    IIf([unitstatus] = "A", 1, 0)
  ) AS Acnt
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN (
      SchGrp
      INNER JOIN LiqPD ON SchGrp.SGrpID = LiqPD.SGrpID
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Units
    INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  DatePart("yyyy", [FundDate_Projected]),
  Clients.ClientShNm,
  MstrLease.MLNo,
  MstrLease.MLBack,
  MstrLease.MLSent,
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.UnitsInGroup,
  Schedule.SchShortDesc,
  Schedule.SchIn,
  Schedule.SchOut,
  SchGrp.SGrpNO,
  SchGrp.OrigEquipCost,
  SchGrp.Assignee_Bank,
  SchGrp.LeaseType
HAVING
  (
    (
      (SchGrp.LeaseType)<> "reporting"
    )
  )
ORDER BY
  Schedule.SchNo;

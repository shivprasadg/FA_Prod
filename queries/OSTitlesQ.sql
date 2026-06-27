SELECT
  Clients.ClientShNm AS Customer,
  Clients.ClientGroupID,
  MstrLease.MLNo AS [Lease NO],
  Schedule.SchNo,
  SchGrp.Assignee_Bank AS Assignee,
  IIf(
    [SchTitleToTrust] = True, "Yes", "No"
  ) AS [Titled to Trust?],
  SchGrp.SGrpNO AS [Group],
  UCase([unitvin]) AS VIN,
  Units.unityr AS [Year],
  Units.unitmake AS Make,
  Units.UnitModel AS Model,
  Units.unittype AS Type,
  Units.UnitSubType AS SubType,
  Units.unittasent AS [App Sent],
  Units.UnitTitleRecd AS [Title Received],
  Units.UnitTitleNumber AS [Title No],
  UCase([UnitStatus]) AS Status,
  Locations.LState AS [Title ST],
  DateDiff(
    "d",
    [unittasent],
    Date()
  ) AS [Days o/s],
  Date() AS [As Of]
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Units
    INNER JOIN Locations ON Units.UnitTitleLocationRef = Locations.LocationID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.unittype)<> "Yard Tractor"
    )
    AND (
      (Units.unittasent) Is Not Null
    )
    AND (
      (Units.UnitTitleNumber) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  )
  OR (
    (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Units.unittype)<> "Yard Tractor"
    )
    AND (
      (Units.UnitTitleNumber) Is Null
    )
    AND (
      (
        UCase([UnitStatus])
      )= "A"
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchID,
  SchGrp.SGrpNO;

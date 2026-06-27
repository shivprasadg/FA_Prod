SELECT
  DatePart("yyyy", [UnitAcceptDate]) AS yr,
  IIf(
    Left([UnitType], 6)= "Daycab",
    "Daycab",
    [unittype]
  ) AS UType,
  Count(Units.UnitID) AS CountOfUnitID,
  IIf(
    IsNull(ClientGroupID),
    [clientcompanyname],
    ClientGroupID
  ) AS CCNm,
  Clients.ClientGroupID,
  Clients.ClientGroupID,
  Units.unittype,
  Units.unitstatus
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
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
GROUP BY
  DatePart("yyyy", [UnitAcceptDate]),
  IIf(
    Left([UnitType], 6)= "Daycab",
    "Daycab",
    [unittype]
  ),
  IIf(
    IsNull(ClientGroupID),
    [clientcompanyname],
    ClientGroupID
  ),
  Clients.ClientGroupID,
  Clients.ClientGroupID,
  Units.unittype,
  Units.unitstatus
HAVING
  (
    (
      (Clients.ClientGroupID) Is Not Null
      And (Clients.ClientGroupID)= forms!MgmtMenu!ClientGroup
    )
    And (
      (Units.unittype)<> "decal"
    )
    And (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "b"
    )
  )
ORDER BY
  DatePart("yyyy", [UnitAcceptDate]),
  IIf(
    Left([UnitType], 6)= "Daycab",
    "Daycab",
    [unittype]
  ),
  Count(Units.UnitID);

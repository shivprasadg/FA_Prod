SELECT
  DatePart("yyyy", [UnitAcceptDate]) AS yr,
  IIf(
    Left([UnitType], 6)= "Daycab",
    "Daycab",
    [unittype]
  ) AS UType,
  Count(Units.unitid) AS CountOfunitid,
  Clients.ClientGroupID,
  Clients.ClientGroupID,
  Clients.clientID,
  Clients.ClientShNm
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
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "c"
      Or (Units.unitstatus)= "tr"
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
  Clients.ClientGroupID,
  Clients.ClientGroupID,
  Clients.clientID,
  Clients.ClientShNm,
  IIf(
    IsNull(ClientGroupID),
    [clientcompanyname],
    ClientGroupID
  )
HAVING
  (
    (
      (Clients.clientID)= forms!MgmtMenu!SelCp
    )
    And (
      (Clients.ClientShNm) Not Like "demo*"
    )
  )
ORDER BY
  DatePart("yyyy", [UnitAcceptDate]),
  IIf(
    Left([UnitType], 6)= "Daycab",
    "Daycab",
    [unittype]
  ),
  Count(Units.unitid);

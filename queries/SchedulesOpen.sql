SELECT
  Clients.ClientShNm AS Lessee,
  Schedule.SchNo,
  Min(Units.UnitActualDelDate) AS EarliestDelvDate,
  Max(Units.UnitActualDelDate) AS LatestDelvDate,
  IIf(
    IsDate([Schout]),
    "Out  " & [SchOut] & Chr(13)& Chr(10),
    Null
  )& IIf(
    IsDate([SchIn]),
    "Back " & [SchIn],
    "Pending"
  ) AS SchOutBack,
  IIf(
    IsDate([SchFEcoRec]),
    [SchFEcoRec],
    "Pending"
  ) AS FinalEconRecv,
  IIf(
    IsDate([SchRALO]),
    "Out " & [SchRALO] & Chr(13)& Chr(10),
    Null
  )& IIf(
    IsDate([SchRALB]),
    "Back  " & [SchRALB],
    "Pending"
  ) AS RALOutBack,
  IIf(
    IsDate([SchOpsPkg]),
    [SchOpsPkg],
    "Pending"
  ) AS OPSPkg,
  IIf(
    IsDate([SchFPkgO]),
    "Out",
    Null
  ) AS PermPkgOut,
  IIf(
    [SchTitleToTrust] = True, "Yes", "No"
  ) AS FATitleTrust,
  Max(
    Nz([UnitIFBank], "n/a")
  ) AS IntermFundBank,
  Max(
    Nz([UnitPayOff], "n/a")
  ) AS IntermFund,
  "Groups:" & Chr(13)& Chr(10)& BLCDInSchedule([SchId]) AS BLCD,
  "Groups:" & Chr(13)& Chr(10)& AssignmentDateInSchedule([SchId]) AS AssignmentDate,
  SchGrp.Assignee_Bank AS Assignee,
  "Groups:" & Chr(13)& Chr(10)& FundDateInSchedule([SchId]) AS FundDate,
  Schedule.SchID,
  Min(SchGrp.SGrpID) AS GroupID,
  Clients.clientID
FROM
  (
    Units
    INNER JOIN (
      Clients
      INNER JOIN (
        (
          SchGrp
          INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
        )
        INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
      ) ON Clients.clientID = MstrLease.Client
    ) ON Units.GroupREF = SchGrp.SGrpID
  )
  LEFT JOIN ScheduleAmends ON Schedule.SchID = ScheduleAmends.ScheduleID
GROUP BY
  Clients.ClientShNm,
  Schedule.SchNo,
  IIf(
    IsDate([Schout]),
    "Out  " & [SchOut] & Chr(13)& Chr(10),
    Null
  )& IIf(
    IsDate([SchIn]),
    "Back " & [SchIn],
    "Pending"
  ),
  IIf(
    IsDate([SchFEcoRec]),
    [SchFEcoRec],
    "Pending"
  ),
  IIf(
    IsDate([SchRALO]),
    "Out " & [SchRALO] & Chr(13)& Chr(10),
    Null
  )& IIf(
    IsDate([SchRALB]),
    "Back  " & [SchRALB],
    "Pending"
  ),
  IIf(
    IsDate([SchOpsPkg]),
    [SchOpsPkg],
    "Pending"
  ),
  IIf(
    IsDate([SchFPkgO]),
    "Out",
    Null
  ),
  IIf(
    [SchTitleToTrust] = True, "Yes", "No"
  ),
  SchGrp.Assignee_Bank,
  Schedule.SchID,
  Clients.clientID,
  MstrLease.MLOrig,
  SchGrp.Assigned_Indicator
HAVING
  (
    (
      (MstrLease.MLOrig)= "FA"
    )
    AND (
      (SchGrp.Assigned_Indicator)<> "F"
    )
  );

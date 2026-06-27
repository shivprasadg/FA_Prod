SELECT
  SchGrp.GroupAddedDate AS DateGroupAdded,
  DatePart('q', [GroupAddedDate]) AS OPSQtr,
  DatePart('yyyy', [GroupAddedDate]) AS OPSYear,
  Nz([sgrpfaid], [schfaid]) AS LFAID,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Group_UnitCount_Cost_Rents.UnitCount,
  Group_UnitCount_Cost_Rents.LORent AS TotalLORent,
  Group_UnitCount_Cost_Rents.SyndRent AS TotalSyndRent,
  Group_UnitCount_Cost_Rents.UnitCost,
  Group_UnitCount_Cost_Rents.ChildPartCost,
  [UnitCost] + [ChildPartCost] AS TotalUnitCost,
  MstrLease.MLID,
  Schedule.SchID,
  SchGrp.SGrpID
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
  INNER JOIN Group_UnitCount_Cost_Rents ON SchGrp.SGrpID = Group_UnitCount_Cost_Rents.GroupID
GROUP BY
  SchGrp.GroupAddedDate,
  DatePart('q', [GroupAddedDate]),
  DatePart('yyyy', [GroupAddedDate]),
  Nz([sgrpfaid], [schfaid]),
  MstrLease.MLNo,
  MstrLease.MLOrig,
  Group_UnitCount_Cost_Rents.UnitCount,
  Group_UnitCount_Cost_Rents.LORent,
  Group_UnitCount_Cost_Rents.SyndRent,
  Group_UnitCount_Cost_Rents.UnitCost,
  Group_UnitCount_Cost_Rents.ChildPartCost,
  [UnitCost] + [ChildPartCost],
  MstrLease.MLID,
  Schedule.SchID,
  SchGrp.SGrpID
HAVING
  (
    (
      (
        Nz([sgrpfaid], [schfaid])
      ) Is Not Null
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  );

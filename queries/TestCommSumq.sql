SELECT
  Salesperson.SPInitials,
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  GrpUnitTypeq.CountOfUnitID AS Expr1,
  GrpUnitTypeq.UnitCost,
  GrpUnitTypeq.AvgOfUnitSynResid AS Expr2,
  GrpUnitTypeq.AvgOfUNITRENT,
  Sum(SGrpExpq.SGExp) AS SumOfSGExp
FROM
  (
    (
      (
        (
          (
            Salesperson
            INNER JOIN Clients ON Salesperson.SalespersonID = Clients.SalespersonREF
          )
          INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
        )
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    LEFT JOIN GrpUnitTypeq ON SchGrp.SGrpID = GrpUnitTypeq.UTGrpRef
  )
  LEFT JOIN SGrpExpq ON SchGrp.SGrpID = SGrpExpq.SGrpID
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
  )
GROUP BY
  Salesperson.SPInitials,
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  GrpUnitTypeq.CountOfUnitID,
  GrpUnitTypeq.UnitCost,
  GrpUnitTypeq.AvgOfUnitSynResid,
  GrpUnitTypeq.AvgOfUNITRENT
HAVING
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
    )
  )
ORDER BY
  Schedule.SchID;

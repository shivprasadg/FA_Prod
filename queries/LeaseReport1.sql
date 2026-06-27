SELECT
  MstrLease.MLNo AS [Lease No],
  MstrLease.MLDt AS MstrLeaseDate,
  MstrLease.MLOrig,
  Schedule.SchFAID AS FAID,
  Schedule.POSentDt AS POSentDate,
  Units.UnitID AS AssetId,
  Units.UnitYr AS ModelYear,
  Units.UnitMake AS Make,
  Units.UnitType,
  Schedule.SchNo,
  SchGrp.SGrpNO AS [Group],
  Units.UnitUnitNum AS [Unit No],
  Units.Unitvin,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & IIf([Axle] = "tbd", Null, " " & [axle]) AS [Desc],
  Clients.clientCompanyName AS Lessee,
  Clients.clientCoNm2 AS Co_Lessee,
  Units.UnitTitleNumber AS [Title Num],
  Units.unitstatus AS Status,
  Units.UnitInServ AS InService,
  CCur(
    Nz([UnitFinalCost], 0)
  ) AS UnitCost,
  Units.UnitInvoiceNum,
  Units.UnitVendorPytDate AS UnitInvoicePaid,
  Units.UnitCOABack,
  Units.UnitEstDelDate,
  SchGrp.FundDate AS SynDate,
  SchGrp.POSentDt AS PODate
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
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Schedule.POSentDt DESC;

SELECT
  Units.UnitTitleLocationRef,
  Units.UnitGarageLocationRef,
  Units.UnitRegLocationRef,
  Units.UnitDeliveryLocationRef,
  Units.UnitType,
  Units.UnitSubType,
  Units.CofASigner,
  MstrLease.MLID,
  Schedule.SchID,
  Units.UnitModel,
  Schedule.SchedDt,
  Units.UnitID,
  Units.unitvin,
  [sgrpno] & [unitdeliverylocationref] & [unitgaragelocationref] & [unittitlelocationref] & [unitregLocationref] & [cofasigner] AS GrpCD,
  Units.UNITMAKE,
  Clients.clientCompanyName,
  Clients.clientCoNm2,
  Clients.clientDBA,
  SchGrp.SGrpID,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  MstrLease.MLNo,
  MstrLease.MLDt,
  DatePart("yyyy", [mldt]) AS MLYr,
  MonthName(
    DatePart("m", [mldt])
  ) AS MLMo,
  DatePart("d", [mldt]) AS mlday,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.PRINTCOFA AS Expr1,
  Right([unitvin], 6) AS vin4,
  Units.unitunitnum,
  [UnitYr] & "   " & [UnitMake] & "   " & [UnitModel] & "   " & " " & [UnitSubtype] & " " & [UnitType] & " " & Nz([AXLE] = "TBD") AS [Desc],
  Units.UnitAcceptDate,
  Mileage.Hours,
  Mileage.MilesLTD AS AquiMI,
  Mileage.Miles AS Mileage,
  Units.Unitstatus,
  SchGrp.SGrpSchID,
  SchGrp.SGRPPLB,
  Units.unitrent,
  SchGrp.BLCD_Projected,
  Schedule.SchFSL,
  SchGrp.SGrpFSLPmt,
  Units.UNITTAX,
  Units.UNITSRENT,
  IIf(
    [unitfinalcost] = 0, Null, [unitfinalcost]
  ) AS Unitcost,
  Units.UnitTaxExempt,
  IIf(
    [SGRPPLB] = True,
    Nz([AcqMileage], 0),
    0
  ) AS PLBMileage,
  SchGrp.SGRPPLB AS PLB
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
    LEFT JOIN Mileage ON Units.UnitID = Mileage.UnitID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (SchGrp.SGrpID)= forms!groupform!sgrpid
    )
    And (
      (Units.Unitstatus)= "B"
    )
  )
ORDER BY
  Clients.clientCompanyName,
  Right([unitvin], 6),
  Units.unitunitnum;

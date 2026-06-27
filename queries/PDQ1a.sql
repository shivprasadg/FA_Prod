SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpID,
  SchGrp.BLCD,
  SchGrp.SGrpNoSlv,
  SchGrp.SGRPNoPD,
  Units.unitunitnum,
  Units.UnitID,
  Units.UnitVIN,
  Units.UnitStatus,
  Units.UnitAcceptDate,
  [unityr] & " " & [unitmake] & " " & " " & [unittype] AS [Desc],
  Units.UnitTaxExempt,
  Locations.LocationName,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  BillStEndq.BillEndDt AS EndDate,
  Locations_1.LAddrLine2 AS BLoST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLoST2,
  Locations_1.LCity AS BLocCity,
  Locations_1.LState AS BLocSt,
  Locations_1.LZip AS BLocZip,
  Clients.clientCompanyName,
  Locations_1.LAddrLine1 AS blOCst2,
  lookup_States.StTaxType,
  Bills.BillNo,
  Clients.FAVenorNo,
  Schedule.SchBnkClNO,
  Right([unitvin], 6) AS vin4,
  Units.unitpdrent,
  BillStEndq.Billstart AS BillStdt,
  Bills.BillID,
  Units.unitinserv,
  Schedule.schcrexp,
  INTDaysQ.SlvDays,
  IIf(
    [sgrpnoslv] = No, 0, [UnitPDRent] * [slvdays]
  ) AS SlvRent,
  IIf(
    [sgrpnoslv] = No,
    0,
    (
      (
        [UnitPDRent] -([SchCurrRent] / 30)
      )
    )
  ) AS DSlvRent,
  IIf([SgrpNoPD] = True, 0, [UnitPDRent]) AS PdRent,
  IIf([Schfsl] = 0, [unittax] / 30, 0) AS Dailytax,
  Locations.LTaxRate,
  Schedule.SchFSL,
  BillUnitq.UnitPDST,
  Units.unitrent,
  Units.unittax,
  Bills.BillDueDt,
  Bills.DueonRec,
  BillStEndq.SlvEnd,
  Bills.Billdt,
  Bills.BillSent,
  INTDaysQ.UBSlvS,
  UCase([invfld1desc] & " " & [invfld1]) AS Bill1,
  Nz([invfld2desc] & " " & [invfld2], "") AS Bill2,
  IIf(
    [unitpdst] > [billenddt],
    0,
    DateDiff(
      "d",
      [unitpdst],
      IIf(
        IsNull([UNITOFfleasedt ]),
        [billenddt],
        [UNITOFFLEASEDT]
      )
    )+ 1
  ) AS PDDAYS,
  Units.unitoffleasedt,
  SchGrp.SGrpFSLPmt,
  IIf([unitpdrent] > 0, 1, 0) AS UnitCount
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        (
          (
            Schedule
            LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
          )
          INNER JOIN Bills ON Schedule.SchID = Bills.BillSchRef
        )
        INNER JOIN BillStEndq ON Schedule.SchID = BillStEndq.SchID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      (
        (
          Locations
          INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
        )
        INNER JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
      )
      INNER JOIN INTDaysQ ON Units.UnitID = INTDaysQ.UnitID
    )
    INNER JOIN BillUnitq ON Units.UnitID = BillUnitq.UnitID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (BillStEndq.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (Units.UnitInServ) Is Not Null
      And (Units.UnitInServ)< Bills.Billto
    )
  )
GROUP BY
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpID,
  SchGrp.BLCD,
  SchGrp.SGrpNoSlv,
  SchGrp.SGRPNoPD,
  Units.unitunitnum,
  Units.UnitID,
  Units.UnitVIN,
  Units.UnitStatus,
  Units.UnitAcceptDate,
  [unityr] & " " & [unitmake] & " " & " " & [unittype],
  Units.UnitTaxExempt,
  Locations.LocationName,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  BillStEndq.BillEndDt,
  Locations_1.LAddrLine2,
  Locations_1.LocationStreet,
  Locations_1.LCity,
  Locations_1.LState,
  Locations_1.LZip,
  Clients.clientCompanyName,
  Locations_1.LAddrLine1,
  lookup_States.StTaxType,
  Bills.BillNo,
  Clients.FAVenorNo,
  Schedule.SchBnkClNO,
  Right([unitvin], 6),
  Units.unitpdrent,
  BillStEndq.Billstart,
  Bills.BillID,
  Units.unitinserv,
  Schedule.schcrexp,
  INTDaysQ.SlvDays,
  IIf(
    [sgrpnoslv] = No, 0, [UnitPDRent] * [slvdays]
  ),
  IIf(
    [sgrpnoslv] = No,
    0,
    (
      (
        [UnitPDRent] -([SchCurrRent] / 30)
      )
    )
  ),
  IIf([SgrpNoPD] = True, 0, [UnitPDRent]),
  IIf([Schfsl] = 0, [unittax] / 30, 0),
  Locations.LTaxRate,
  Schedule.SchFSL,
  BillUnitq.UnitPDST,
  Units.unitrent,
  Units.unittax,
  Bills.BillDueDt,
  Bills.DueonRec,
  BillStEndq.SlvEnd,
  Bills.Billdt,
  Bills.BillSent,
  INTDaysQ.UBSlvS,
  UCase([invfld1desc] & " " & [invfld1]),
  Nz([invfld2desc] & " " & [invfld2], ""),
  IIf(
    [unitpdst] > [billenddt],
    0,
    DateDiff(
      "d",
      [unitpdst],
      IIf(
        IsNull([UNITOFfleasedt ]),
        [billenddt],
        [UNITOFFLEASEDT]
      )
    )+ 1
  ),
  Units.unitoffleasedt,
  SchGrp.SGrpFSLPmt,
  IIf([unitpdrent] > 0, 1, 0),
  Locations_1.LAddrLine1,
  Bills.Billdt,
  SchGrp.SGrpFSLPmt
HAVING
  (
    (
      (Bills.BillID)= [Forms]![Billing]![BillingSubForm].[Form]![BillID]
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Units.unitunitnum,
  Locations.LocationName,
  Right([unitvin], 6);

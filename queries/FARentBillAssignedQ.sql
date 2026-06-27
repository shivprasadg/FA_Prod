SELECT
  Bills.BillID,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  SchGrp.Assigned_Indicator,
  SchGrp.LoanBank,
  SchGrp.SGrpFABill,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  [UnitYr] & " " & [UnitMake] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ) AS [Desc],
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Units.UnitAcceptDate,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2 AS BLocST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLocST2,
  Locations_1.LCity AS BLocCity,
  Locations_1.LState AS BLocSt,
  Locations_1.LZip AS BLocZip,
  Clients.clientCompanyName,
  lookup_States.StTaxType,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Right([unitvin], 6) AS vin4,
  SchGrp.SGrpFSLPmt,
  Nz([SGrpFSLPmt], 0) AS FSLFee,
  IIf(
    [unitoffleasedt] < [billfrom],
    0,
    IIf(
      [Forms]![Billing]![BillingSubForm].[Form]![UseSchRent] =-1,
      [UnitSRent],
      IIf(
        [unitrent] = 0, [UnitSRent], [UnitRent]
      )
    )
  ) AS UnitsBillRent,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  IIf(
    [UnitTaxExempt] = True, 0, [Locations].[LTaxRate] * [UnitsBillRent]
  ) AS TaxAmt,
  Locations.LTaxRate,
  Units.unittax,
  Clients.FAVenorNo,
  Units.unitoffleasedt,
  [invfld1desc] & " " & [invfld1] AS Bill1,
  [invfld2desc] & " " & [invfld2] & " " AS Bill2,
  Bills.DueonRec,
  SchGrp.AssignmentDate,
  Schedule.invfld1desc,
  Schedule.invfld1,
  Schedule.invfld2desc,
  Schedule.invfld2,
  Clients.clientID,
  [UnitsbillRent] AS RentNoTax,
  IIf(
    [UnitTaxExempt] = False,
    Nz([FSLFee], 0)* Nz([locations].[LTaxRate], 0),
    0
  ) AS FSLTax,
  [UnitsBillRent] + Nz([FSLFee], 0)+ Nz([TaxAmt], 0) AS UnitTotalRent,
  [BLocStreet] & Chr(13)& Chr(10)& [BLocSt2] & Chr(13)& Chr(10)& IIf(
    IsNull([BLocST3]),
    Null,
    [BLocST3] & Chr(13)& Chr(10)
  )& [blOCcITY] & ", " & [bLOCST] & " " & Left([BlOCzIP], 5)& "-" & Right([BlOCzIP], 4) AS BillToAddress,
  Bills.BillNo AS InvoiceNumber,
  IIf(
    [DueonRec] =-1,
    "On Receipt",
    IIf(
      IsNull([BillDueDt]),
      [billFrom],
      [BillDueDt]
    )
  ) AS DueDate,
  [SchBnkClNO] & IIf(
    Nz([SgrpBnkNo], "")= "",
    Null,
    "-" & [SgrpBnkNo]
  ) AS BankContract,
  SchGrp.SGrpID AS GroupID
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        (
          Schedule
          LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
        )
        LEFT JOIN Bills ON Schedule.SchID = Bills.BillSchRef
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Locations
      INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
    )
    INNER JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (Bills.Billto)<= Nz([unitoffleasedt], #1/1/2400#)
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

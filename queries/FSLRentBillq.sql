SELECT
  Bills.BillID,
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
  [UnitFinalCost] + IIf(
    IsNull([sumofchildfinalcost]),
    0,
    [sumofchildfinalcost]
  ) AS UCost,
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Units.UnitAcceptDate,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2 AS BLoST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLoST2,
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
  [Unittax] + [unitsrent] + [SGrpFslPMt] AS [UNITS RENT],
  Units.UnitTax,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  IIf(
    [unittaxexempt] =-1, 0, [locations].[LTaxRate] * [unitsrent]
  ) AS TaxAmt,
  Locations.LTaxRate,
  Clients.FAVenorNo,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  SchGrp.Assigned_Indicator,
  SchGrp.LoanBank,
  SchGrp.SGrpFABill,
  Units.unitoffleasedt,
  UCase([invfld1desc] & " " & [invfld1]) AS Bill1,
  UCase(
    [invfld2desc] & " " & [invfld2] & " " & [lesseeacct]
  ) AS Bill2
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
            INNER JOIN PDstdtq ON Schedule.SchID = PDstdtq.SchID
          )
          LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
        )
        LEFT JOIN Bills ON Schedule.SchID = Bills.BillSchRef
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      (
        Locations
        INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
      )
      INNER JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
    )
    LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Bills.BillID)= Forms!Billing!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Null
    )
    And (
      (SchGrp.LoanBank) Is Null
    )
  )
  Or (
    (
      (Bills.BillID)= Forms!Billing!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Not Null
    )
  )
  Or (
    (
      (Bills.BillID)= Forms!Billing!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Null
    )
    And (
      (SchGrp.LoanBank) Is Not Null
    )
    And (
      (SchGrp.SGrpFABill)= True
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

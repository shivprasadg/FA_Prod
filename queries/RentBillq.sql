SELECT
  Bills.BillID,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [unittype] AS [Desc],
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
  IIf(
    (
      [unitstatus] = "O"
      Or [unitstatus] = "O"
      Or [unitstatus] = "S"
      Or [unitstatus] = "r"
    )
    And [unitoffleasedt] < [billfrom],
    0,
    [unitrent]
  ) AS UnitsbillRent,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  IIf(
    [UnitTaxExempt] = True,
    0,
    IIf(
      (
        [unitstatus] = "O"
        Or [unitstatus] = "r"
        Or [unitstatus] = "s"
      )
      And [unitoffleasedt] < [billfrom],
      0,
      [locations].[LTaxRate] * [unitrent]
    )
  ) AS TaxAmt,
  Locations.LTaxRate,
  Units.unittax,
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
  ) AS Bill2,
  Bills.DueonRec
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
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Null
      Or (SchGrp.Assignee_Bank)= "FA llc"
    )
    And (
      (SchGrp.LoanBank) Is Null
    )
  )
  Or (
    (
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Not Null
    )
    And (
      (SchGrp.Assigned_Indicator)= "p"
    )
  )
  Or (
    (
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (SchGrp.Assignee_Bank) Is Null
      Or (SchGrp.Assignee_Bank)= "FA llc"
    )
    And (
      (SchGrp.LoanBank) Is Not Null
    )
    And (
      (SchGrp.SGrpFABill)= Yes
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

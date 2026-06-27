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
  IIf(
    [UNITSRENT] > 0,
    IIf(
      [ScheduleGroupPaymentTerm] > 0,
      IIf(
        [Billto] < DateAdd(
          "m", [ScheduleGroupPaymentTerm],
          [BLCD]
        ),
        [SchGrp].[SGrpFSLPmt],
        0
      ),
      [SchGrp].[SGrpFSLPmt]
    ),
    0
  ) AS FSLFee,
  SchGrp.SGrpFSLPmt,
  Units.UNITSRENT,
  Units.UnitTax,
  IIf(
    [unitstatus] = "A"
    Or IsNull([unitoffleasedt])
    Or [unitoffleasedt] < [billfrom],
    [unitrent],
    0
  ) AS UnitsBillRent,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  IIf(
    [unittaxexempt] =-1,
    0,
    Round(
      [locations].[LTaxRate] * [unitsrent],
      2
    )
  ) AS TaxAmt,
  Locations.LTaxRate,
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
  Locations_1.LAddrLine1 AS BLocSt2,
  Schedule.SchBnkClNO,
  Clients.clientID,
  SchGrp.SgrpBnkNo,
  Locations_1.LAddrLine1,
  Locations_1.LAddrLine2,
  Locations_1.LAddrLine3,
  Sum(BillAdjT.BillAdj) AS SumOfBillAdj,
  DateAdd(
    "m", [ScheduleGroupPaymentTerm],
    [BLCD]
  ) AS DateFeesEnd,
  IIf(
    [Billto] < [DateFeesEnd], "Yes", "No"
  ) AS ChargeFee,
  IIf(
    (
      [unitstatus] = "O"
      Or [unitstatus] = "S"
      Or [unitstatus] = "R"
    )
    And [unitoffleasedt] < [billfrom],
    0,
    [unitsrent]
  ) AS LineTot,
  Bank_Assignee.ShortName,
  Bank_Assignee.BankName AS Remitline1,
  Bank_Assignee.LAddrLine1 AS Remitline2,
  Bank_Assignee.LAddrLine2 AS Remitline3,
  [Bank_Assignee].[LCity] & " " & [Bank_Assignee].[LState] & "  " & Left([Bank_Assignee].[LZip], 5) AS Remitline4,
  IIf([UNITSRENT] > 0, 1, 0) AS UnitCount,
  SchGrp.sgrpid
FROM
  (
    (
      (
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
      )
      LEFT JOIN BillAdjT ON Bills.BillID = BillAdjT.BillAdjBillref
    )
    INNER JOIN LocationsRemit ON MstrLease.Client = LocationsRemit.ClientREF
  )
  LEFT JOIN Bank_Assignee ON SchGrp.Assignee_Bank = Bank_Assignee.ShortName
GROUP BY
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
  ),
  [UnitFinalCost] + IIf(
    IsNull([sumofchildfinalcost]),
    0,
    [sumofchildfinalcost]
  ),
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Units.UnitAcceptDate,
  SchGrp.BLCD_Projected,
  Locations_1.LocationStreet,
  Locations_1.LCity,
  Locations_1.LState,
  Locations_1.LZip,
  Clients.clientCompanyName,
  lookup_States.StTaxType,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Right([unitvin], 6),
  SchGrp.SGrpFSLPmt,
  Units.UNITSRENT,
  Units.UnitTax,
  IIf(
    [unitstatus] = "A"
    Or IsNull([unitoffleasedt])
    Or [unitoffleasedt] < [billfrom],
    [unitrent],
    0
  ),
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  IIf(
    [unittaxexempt] =-1,
    0,
    Round(
      [locations].[LTaxRate] * [unitsrent],
      2
    )
  ),
  Locations.LTaxRate,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  SchGrp.Assigned_Indicator,
  SchGrp.LoanBank,
  SchGrp.SGrpFABill,
  Units.unitoffleasedt,
  UCase([invfld1desc] & " " & [invfld1]),
  UCase(
    [invfld2desc] & " " & [invfld2] & " " & [lesseeacct]
  ),
  Schedule.SchBnkClNO,
  Clients.clientID,
  SchGrp.SgrpBnkNo,
  Locations_1.LAddrLine1,
  Locations_1.LAddrLine2,
  Locations_1.LAddrLine3,
  DateAdd(
    "m", [ScheduleGroupPaymentTerm],
    [BLCD]
  ),
  IIf(
    (
      [unitstatus] = "O"
      Or [unitstatus] = "S"
      Or [unitstatus] = "R"
    )
    And [unitoffleasedt] < [billfrom],
    0,
    [unitsrent]
  ),
  Bank_Assignee.ShortName,
  Bank_Assignee.BankName,
  Bank_Assignee.LAddrLine1,
  Bank_Assignee.LAddrLine2,
  [Bank_Assignee].[LCity] & " " & [Bank_Assignee].[LState] & "  " & Left([Bank_Assignee].[LZip], 5),
  IIf([UNITSRENT] > 0, 1, 0),
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitType] & " " & [UnitSubType],
  IIf(
    [UNITSRENT] > 0,
    IIf(
      [ScheduleGroupPaymentTerm] > 0,
      IIf(
        [Billto] < DateAdd(
          "m", [ScheduleGroupPaymentTerm],
          SchGrp.BLCD
        ),
        [SchGrp].[SGrpFSLPmt],
        0
      ),
      [SchGrp].[SGrpFSLPmt]
    ),
    0
  ),
  Clients.FAVenorNo,
  Locations_1.LAddrLine1,
  Locations_1.LAddrLine1,
  Locations_1.LAddrLine2,
  SchGrp.sgrpid
HAVING
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

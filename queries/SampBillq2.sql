SELECT
  Schedule.SchID,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS [Desc],
  Units.UnitFinalCost,
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState AS St,
  Locations.LZip,
  Units.UnitAcceptDate,
  Units.UnitPDStartDate,
  SchGrp.BLCD,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2 AS BLoST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLoST2,
  Locations_1.LCity AS BLocCity,
  Locations_1.LState AS BLocSt,
  Locations_1.LZip AS BLocZip,
  Clients.clientCompanyName,
  Locations_1.LAddrLine1 AS blOCst2,
  lookup_States.StTaxType,
  Clients.FAVenorNo,
  Schedule.SchBnkClNO,
  Right([unitvin], 6) AS vin4,
  Units.UnitAcceptDate,
  SchGrp.SGrpFSLPmt,
  IIf(
    [unitstatus] = "O"
    Or [unitstatus] = "r"
    Or [unitstatus] = "s",
    0, [unitsrent]
  ) AS URenT,
  Units.UnitTax,
  Clients_1.clientCompanyName,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SgrpBnkNo,
  Schedule.SchBnkClNO,
  Clients.ClientShNm,
  IIf(
    [unitstatus] = "o"
    Or [unitstatus] = "r"
    Or [unitstatus] = "s",
    0,
    IIf(
      [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
    )
  ) AS TaxAmt,
  Locations.LTaxRate,
  Locations.LState AS LocSt,
  Units.UNITSRENT,
  Clients_1.clientAddress1,
  Clients_1.clientAddress2,
  Clients_1.clientCity,
  Clients_1.clientState,
  Clients_1.clientZip,
  Units.unitstatus,
  IIf(
    IsNull([AssignmentDate]),
    [BLCD],
    [AssignmentDate]
  ) AS 1BankRent,
  Locations_1.LAddrLine2 AS bloc3,
  BankBillAddq.LocationStreet,
  BankBillAddq.LAddrLine1,
  BankBillAddq.LAddrLine2,
  BankBillAddq.LAddrLine3,
  BankBillAddq.LCity,
  BankBillAddq.LState,
  BankBillAddq.LZip
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
          LEFT JOIN PDstdtq ON Schedule.SchID = PDstdtq.SchID
        )
        LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN (
      (
        SchGrp
        LEFT JOIN Clients AS Clients_1 ON SchGrp.Assignee_Bank = Clients_1.ClientShNm
      )
      LEFT JOIN BankBillAddq ON Clients_1.clientID = BankBillAddq.clientID
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Locations
      INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
    )
    LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (MstrLease.MLNo)= forms!Billing!SelLease
    )
    And (
      (Schedule.SchID)= forms!Billing!selsch
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (Schedule.SchFSL)= No
    )
  )
ORDER BY
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

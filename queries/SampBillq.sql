SELECT
  Clients_1.clientID,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & IIf(
    [UnitType] = "Tractor", '', " " & [UnitType]
  ) AS [Desc],
  Units.UnitFinalCost,
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations_2.LocationType,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Units.UnitPDStartDate,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2 AS BLoST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLoST2,
  Locations_1.LCity AS BLocCity,
  Locations_1.LState AS BLocSt,
  Locations_1.LZip AS BLocZip,
  Locations_1.LAddrLine1 AS blOCst2,
  lookup_States.StTaxType,
  Schedule.SchBnkClNO,
  Right([unitvin], 6) AS vin4,
  Units.UnitAcceptDate,
  SchGrp.SGrpFSLPmt,
  Units.UNITSRENT,
  Units.UnitTax,
  Clients_1.clientCompanyName,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SgrpBnkNo,
  Schedule.SchBnkClNO,
  Locations.LTaxRate,
  IIf(
    [unittaxexempt] =-1,
    0,
    [locations].[LTaxRate] * IIf(
      (
        [unitstatus] = "O"
        Or [unitstatus] = "O"
        Or [unitstatus] = "S"
        Or [unitstatus] = "r"
      )
      And [unitoffleasedt] < [billfrom],
      0,
      [unitsrent]
    )
  ) AS TaxAmt,
  Units.UNITSRENT,
  IIf(
    [unittaxexempt] =-1,
    0,
    [locations].[LTaxRate] * IIf(
      (
        [unitstatus] = "O"
        Or [unitstatus] = "O"
        Or [unitstatus] = "S"
        Or [unitstatus] = "r"
      )
      And [unitoffleasedt] < [billfrom],
      0,
      [unitsrent]
    )
  )+ IIf(
    (
      [unitstatus] = "O"
      Or [unitstatus] = "O"
      Or [unitstatus] = "S"
      Or [unitstatus] = "r"
    )
    And [unitoffleasedt] < [billfrom],
    0,
    [unitsrent]
  )+ Nz([SGrpFSLPmt], 0) AS LineTot,
  Clients_1.clientAddress1,
  Clients_1.clientAddress2,
  Clients_1.clientCity,
  Clients_1.clientState,
  Clients_1.clientZip,
  Locations_2.LocationStreet,
  Locations_2.LAddrLine1,
  Locations_2.LAddrLine2,
  Locations_2.LCity,
  Locations_2.LState,
  Locations_2.LZip,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Bills.BillID,
  PDstdtq.MLNo,
  Sum(BillAdjT.BillAdj) AS SumOfBillAdj,
  Locations_2.LAddrLine3,
  IIf(
    [UnitOffLeasedt] < [BillFrom], 0, [UnitRent]
  ) AS UnitsBillRent,
  Units.UCCC
FROM
  (
    (
      (
        (
          (
            Schedule
            LEFT JOIN PDstdtq ON Schedule.SchID = PDstdtq.SchID
          )
          LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
        )
        INNER JOIN Bills ON Schedule.SchID = Bills.BillSchRef
      )
      LEFT JOIN BillAdjT ON Bills.BillID = BillAdjT.BillAdjBillref
    )
    INNER JOIN (
      (
        SchGrp
        INNER JOIN Clients AS Clients_1 ON SchGrp.Assignee_Bank = Clients_1.ClientShNm
      )
      INNER JOIN Locations AS Locations_2 ON Clients_1.clientID = Locations_2.ClientREF
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Locations
      INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
    )
    LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  Clients_1.clientID,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & IIf(
    [UnitType] = "Tractor", '', " " & [UnitType]
  ),
  Units.UnitFinalCost,
  Units.UnitTaxExempt,
  Locations.LocationStreet,
  Locations_2.LocationType,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Units.UnitPDStartDate,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2,
  Locations_1.LocationStreet,
  Locations_1.LCity,
  Locations_1.LState,
  Locations_1.LZip,
  Locations_1.LAddrLine1,
  lookup_States.StTaxType,
  Right([unitvin], 6),
  Units.UnitAcceptDate,
  SchGrp.SGrpFSLPmt,
  Units.UnitTax,
  Clients_1.clientCompanyName,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SgrpBnkNo,
  Schedule.SchBnkClNO,
  Locations.LTaxRate,
  IIf(
    [unittaxexempt] =-1,
    0,
    [locations].[LTaxRate] * IIf(
      (
        [unitstatus] = "O"
        Or [unitstatus] = "O"
        Or [unitstatus] = "S"
        Or [unitstatus] = "r"
      )
      And [unitoffleasedt] < [billfrom],
      0,
      [unitsrent]
    )
  ),
  Units.UNITSRENT,
  IIf(
    [unittaxexempt] =-1,
    0,
    [locations].[LTaxRate] * IIf(
      (
        [unitstatus] = "O"
        Or [unitstatus] = "O"
        Or [unitstatus] = "S"
        Or [unitstatus] = "r"
      )
      And [unitoffleasedt] < [billfrom],
      0,
      [unitsrent]
    )
  )+ IIf(
    (
      [unitstatus] = "O"
      Or [unitstatus] = "O"
      Or [unitstatus] = "S"
      Or [unitstatus] = "r"
    )
    And [unitoffleasedt] < [billfrom],
    0,
    [unitsrent]
  )+ Nz([SGrpFSLPmt], 0),
  Clients_1.clientAddress1,
  Clients_1.clientAddress2,
  Clients_1.clientCity,
  Clients_1.clientState,
  Clients_1.clientZip,
  Locations_2.LocationStreet,
  Locations_2.LAddrLine1,
  Locations_2.LAddrLine2,
  Locations_2.LCity,
  Locations_2.LState,
  Locations_2.LZip,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Bills.BillID,
  PDstdtq.MLNo,
  Locations_2.LAddrLine3,
  IIf(
    [UnitOffLeasedt] < [BillFrom], 0, [UnitRent]
  ),
  Units.UCCC,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitType] & " " & [UnitSubType],
  Locations_1.LAddrLine1,
  Schedule.SchBnkClNO,
  Units.UNITSRENT,
  Locations_2.LocObsolete,
  SchGrp.BLCD
HAVING
  (
    (
      (Locations_2.LocationType)= "billing"
    )
    And (
      (Bills.BillID)= Forms!Billing!BillID
    )
    And (
      (Locations_2.LocObsolete)= No
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
  )
ORDER BY
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

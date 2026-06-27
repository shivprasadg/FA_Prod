SELECT
  Clients_1.clientID,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Locations.LocationName,
  Units.unitunitnum,
  Units.UnitStatus,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS [Desc],
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
  Units.UNITSRENT AS UnitsBillRent,
  Units.UNITSRENT,
  Units.UnitTax,
  Clients_1.clientCompanyName AS aSSIGNEEnM,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SgrpBnkNo,
  Schedule.SchBnkClNO,
  Locations.LTaxRate,
  IIf(
    [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
  ) AS TaxAmt,
  Units.UNITSRENT,
  IIf(
    [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
  )+ [unitsrent] AS LineTot,
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
  Locations_2.LocationType,
  Locations_2.LZip,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Bills.BillID,
  PDstdtq.MLNo,
  Sum(BillAdjT.BillAdj) AS SumOfBillAdj,
  Locations_2.LocObsolete,
  Locations_2.LAddrLine3,
  Units.unitvin,
  Units.unitmake,
  Units.unittype,
  SchGrp.AssignmentDate,
  Schedule.schbnkclno,
  IIf([UNITSRENT] > 0, 1, 0) AS UnitCount
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
  Units.UnitStatus,
  Units.UnitFinalCost,
  Units.UnitTaxExempt,
  Locations.LocationStreet,
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
    [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
  ),
  Units.UNITSRENT,
  IIf(
    [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
  )+ [unitsrent],
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
  Locations_2.LocationType,
  Locations_2.LZip,
  Bills.BillNo,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  Bills.BillID,
  PDstdtq.MLNo,
  Locations_2.LocObsolete,
  Locations_2.LAddrLine3,
  Units.unitvin,
  Units.unitmake,
  Units.unittype,
  SchGrp.AssignmentDate,
  Schedule.schbnkclno,
  IIf([UNITSRENT] > 0, 1, 0),
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitType] & " " & [UnitSubType],
  Locations_1.LAddrLine1,
  Units.UNITSRENT,
  Units.UNITSRENT,
  Locations_2.LocationType,
  Schedule.schbnkclno,
  SchGrp.BLCD
HAVING
  (
    (
      (Units.UnitStatus)= "a"
    )
    And (
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (Locations_2.LocObsolete)= No
    )
    And (
      (Locations_2.LocationType)= "billing"
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
  )
ORDER BY
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

SELECT
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.AssignmentDate,
  Schedule.SchFSL,
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
  Locations.LState,
  Locations.LZip,
  Units.UnitAcceptDate,
  Units.UnitPDStartDate,
  SchGrp.BLCD_Projected,
  Locations_1.LAddrLine2 AS BLoST3,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLoST2,
  Locations_1.LCity AS BLocCity,
  Locations_1.LState AS BLocSt,
  Locations_1.LZip AS BLocZip,
  Clients.clientCompanyName,
  Locations_2.LocationType,
  Locations_1.LAddrLine1 AS blOCst2,
  lookup_States.StTaxType,
  Clients.FAVenorNo,
  Schedule.SchPDInv,
  Schedule.SchPDInvDt,
  Schedule.SchPDBillDt,
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
  Clients.ClientShNm,
  IIf(
    [unittaxexempt] = True, 0, [locations].[LTaxRate] * [unitsrent]
  ) AS TaxAmt,
  Locations.LTaxRate,
  Units.UNITSRENT,
  Clients_1.clientAddress1,
  Clients_1.clientAddress2,
  Clients_1.clientCity,
  Clients_1.clientState,
  Clients_1.clientZip,
  Locations_2.LocationType,
  Locations_2.LocationStreet,
  Locations_2.LAddrLine1,
  Locations_2.LAddrLine2,
  Locations_2.LCity,
  Locations_2.LState,
  Locations_2.LZip,
  IIf(
    IsNull([AssignmentDate]),
    [BLCD],
    [AssignmentDate]
  ) AS 1stBankRnt,
  [LeaseTermPrimary] + DateDiff(
    "m",
    IIf(
      IsNull([AssignmentDate]),
      [BLCD],
      [AssignmentDate]
    ),
    [BLCD]
  ) AS [No Pmts],
  Units.unitstatus,
  IIf(
    [unitstatus] = "o"
    Or [unitstatus] = "s"
    Or [unitstatus] = "r",
    0,
    IIf(
      IsNull([unitsrent]),
      [unitrent],
      [unitsrent]
    )
  ) AS FSLrent,
  IIf(
    [unitstatus] = "o"
    Or [unitstatus] = "r"
    Or [unitstatus] = "s",
    0, [SgrpFSLPmt]
  ) AS FSLServ,
  Units.unittax,
  Units.unitrent
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        Schedule
        LEFT JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN (
      (
        SchGrp
        LEFT JOIN Clients AS Clients_1 ON SchGrp.Assignee_Bank = Clients_1.ClientShNm
      )
      LEFT JOIN Locations AS Locations_2 ON Clients_1.clientID = Locations_2.ClientREF
    ) ON Schedule.SchID = SchGrp.SGrpSchID
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
      (MstrLease.MLNo)= forms!Billing!SELLEASE
    )
    And (
      (Schedule.SchID)= forms!Billing!SELsCH
    )
    And (
      (Schedule.SchFSL)= Yes
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (Locations_2.LocationType)= "billing"
    )
    And (
      (SchGrp.Assignee_Bank) Is Not Null
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID,
  Locations.LocationName,
  Units.unitunitnum,
  Right([unitvin], 6);

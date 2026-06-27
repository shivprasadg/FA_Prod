SELECT
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.BLCD,
  IIf(
    [UnitType] = "Tractor",
    Round(0.0329 * [sgrpfslpmt], 2),
    0
  ) AS DailyServ,
  IIf(
    [UnitType] = "Tractor",
    Round(0.0329 * [sgrprsltpmt], 2),
    0
  ) AS [DailyNon-Serv],
  Units.UnitInServ,
  Units.UnitDocAccpt,
  Units.UnitInServInd,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Locations_1.LocationStreet AS BLocStreet,
  Locations_1.LAddrLine1 AS BLOCst2,
  Locations_1.LAddrLine2 AS BLOCST3,
  Locations_1.LCity AS BLOCCity,
  Locations_1.LState AS BLOCST,
  Locations.LZip AS BLOCZip,
  Clients.FAVenorNo,
  UCase([invfld1desc] & " " & [invfld1]) AS Bill1,
  UCase(
    [invfld2desc] & " " & IIf([invfld2] = "0", Null, [invfld2])
  ) AS Bill2,
  Units.unitunitnum,
  SchGrp.SGRPNoPD,
  Units.UnitTaxExempt,
  Units.unitmake,
  MstrLease.MLNo,
  SchGrp.SGrpNO,
  Locations.LocationID,
  Locations.LocationName,
  Locations.LTaxRate,
  lookup_States.StTaxType,
  [Forms]![Billing]![BillingSubForm].[Form]![BillNo] AS InvoiceNumber,
  [Forms]![Billing]![BillingSubForm].[Form]![billfrom] AS invfrdt,
  [Forms]![Billing]![BillingSubForm].[Form]![billto] AS invtodt,
  [Forms]![Billing]![BillingSubForm].[Form]![billduedt] AS invduedt,
  [Forms]![Billing]![BillingSubForm].[Form]![billdt] AS invdt,
  Schedule.SchFSL,
  SchGrp.sgrprsltpmt,
  Units.unitvin,
  Units.UnitMake,
  Units.UnitType,
  Units.unitpdrent,
  SchGrp.SGrpID
FROM
  (
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
      (
        Units
        LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
      )
      LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN Locations AS Locations_1 ON Clients.ClientGroupId = Locations_1.ClientGroupID
WHERE
  (
    (
      (Schedule.SchID)= Forms!Billing!SchID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (Units.UnitInServ) Is Not Null
    )
    And (
      (Locations_1.LocationType)= "billing"
    )
  );

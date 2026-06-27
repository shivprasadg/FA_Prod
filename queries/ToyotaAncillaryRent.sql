SELECT
  Bills.BillID,
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.BLCD,
  [locations_1].[LocationStreet] & Chr(13)& Chr(10)& IIf(
    IsNull([locations_1].[LAddrLine1]),
    Null,
    [locations_1].[LAddrLine1] & Chr(13)& Chr(10)
  )& IIf(
    IsNull([locations_1].[LAddrLine2]),
    Null,
    [locations_1].[LAddrLine2] & Chr(13)& Chr(10)
  )& [locations_1].[LCity] & " " & [locations_1].[LState] & " " & [locations_1].[LZip] AS BillingAddress,
  IIf(
    [UnitType] = "Tractor", [Sgrpfslpmt],
    0
  ) AS Serv,
  IIf(
    [UnitType] = "Tractor", [sgrprsltpmt],
    0
  ) AS [Non-Serv],
  Units.UnitInServ,
  Units.UnitInServInd,
  Clients.FAVenorNo,
  UCase([invfld1desc] & " " & [invfld1]) AS Bill1,
  UCase(
    [invfld2desc] & " " & IIf([invfld2] = "0", Null, [invfld2])
  ) AS Bill2,
  Units.unitunitnum,
  Units.UnitTaxExempt,
  Units.unitmake,
  MstrLease.MLNo,
  SchGrp.SGrpNO,
  Locations.LocationID,
  Locations.LocationName,
  Locations.LState,
  Locations.LCounty,
  Locations.LCity,
  Locations.LTaxRate,
  lookup_States.StTaxType,
  Bills.billno AS InvoiceNumber,
  Bills.Billfrom AS invfrdt,
  Bills.Billto AS invtodt,
  Bills.BillDueDt AS invduedt,
  Bills.Billdt AS invdt,
  SchGrp.SGrpFSLPmt,
  SchGrp.sgrprsltpmt,
  Units.UNITRENT,
  Units.unitmodel,
  Units.UnitVIN,
  SchGrp.AssignmentDate,
  Banks.BankName,
  Units.unitoffleasedt,
  IIf(
    [UnitRent] < 1,
    0,
    [serv] + [non-serv] + Round(
      IIf(
        [UnitTaxExempt] = Yes, 0, [non-serv] * [Locations].[LTaxRate]
      ),
      2
    )
  ) AS UnitLineTotal,
  Units.UnitType,
  Units.UnitSubType
FROM
  (
    (
      Bills
      INNER JOIN (
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
      ) ON Bills.BillSchRef = Schedule.SchID
    )
    INNER JOIN Locations AS Locations_1 ON SchGrp.BillToAddressID = Locations_1.LocationID
  )
  INNER JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
WHERE
  (
    (
      (Bills.BillID)= Forms!Billing!BillID
    )
    And (
      (SchGrp.BLCD) Is Not Null
    )
    And (
      (Units.unitoffleasedt) Is Null
      Or (Units.unitoffleasedt)>= bills.Billfrom
    )
  )
ORDER BY
  Units.unitunitnum;

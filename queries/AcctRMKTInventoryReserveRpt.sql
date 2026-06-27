SELECT
  DISTINCTROW vw_Sch_Grp_Units_Remarketing.ClientShortName AS Client,
  vw_Sch_Grp_Units_Remarketing.AssetID,
  vw_Sch_Grp_Units_Remarketing.UnitNum,
  vw_Sch_Grp_Units_Remarketing.VIN,
  [MYear] & " " & [Make] & "  " & [Model] & " " & [Type] & " " & [axle] AS Description,
  Locations.LocationName AS Location,
  [LAddrLine1] & " " & [LCity] & " " & [LState] & " " & [LZip] AS [Location Address],
  CCur(
    Nz([Rmkt].[BuyOutCost], 0)
  ) AS [BuyOut Cost],
  Rmkt.InvoicePaidDate,
  Rmkt.BuyOutDate,
  vw_Sch_Grp_Units_Remarketing.Status,
  Banks.BankShortName AS InterimBank,
  CCur(
    Nz([Rmkt.IFAmount], 0)
  ) AS InterimAmt,
  Rmkt.IFLoanNumber AS AccountNumber,
  Rmkt.IFDate AS InterimDate,
  Rmkt.IFPaidDate AS InterimPaidDate,
  CCur(
    Nz([Rmkt].[BuyOutCost], 0)
  )-(
    Nz([Rmkt].[IFAmount], 0)
  ) AS FAAmount,
  IIf(
    IsNull([Rmkt].[IFAmount]),
    Null,
    Round(
      1 -(
        [Rmkt].[IFAmount] / [Rmkt].[BuyOutCost]
      ),
      3
    )
  ) AS FAPercent,
  Clients.ClientShNm AS Buyer,
  CCur(
    Nz([Rmkt].[SoldPrice], 0)
  ) AS SoldPrice,
  IIf(
    Nz([Rmkt].[SoldPrice], 0)< Nz([Rmkt].[BuyOutCost], 0),
    CCur(
      Nz([Rmkt].[SoldPrice], 0)- Nz([Rmkt].[BuyOutCost], 0)
    ),
    0
  ) AS Reserve,
  IIf(
    Nz([Rmkt].[SoldPrice], 0)< Nz([Rmkt].[BuyOutCost], 0),
    0,
    CCur(
      Nz([Rmkt].[SoldPrice], 0)- Nz([Rmkt].[BuyOutCost], 0)
    )
  ) AS Profit
FROM
  (
    (
      (
        vw_Sch_Grp_Units_Remarketing
        INNER JOIN (
          Units
          LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
        ) ON vw_Sch_Grp_Units_Remarketing.AssetID = Units.UnitID
      )
      INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    )
    LEFT JOIN Banks ON Rmkt.IFBankID = Banks.BankId
  )
  LEFT JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Between [forms]![acctmenu]![StartDate]
      And [forms]![acctmenu]![EndDate]
    )
  )
ORDER BY
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  vw_Sch_Grp_Units_Remarketing.AssetID,
  vw_Sch_Grp_Units_Remarketing.UnitNum;

SELECT
  DISTINCT vw_Sch_Grp_Units_Remarketing.ClientShortName AS Client,
  vw_Sch_Grp_Units_Remarketing.AssetID,
  vw_Sch_Grp_Units_Remarketing.UnitNum,
  vw_Sch_Grp_Units_Remarketing.VIN,
  LookUp_UnitDesc.DescriptionMed AS Description,
  Locations.LocationName AS Location,
  [LAddrLine1] & " " & [LCity] & " " & [LState] & " " & [LZip] AS [Location Address],
  CCur(
    Nz([Rmkt].[BuyOutCost], 0)
  ) AS [BuyOut Cost],
  CCur(
    IIf(
      [vw_Sch_Grp_Units_Remarketing].[ClientGroupId] = 432
      And [rmkt].[BuyOutDate] > #6/15/2017#
      And [rmkt].[BuyOutDate] < #7/5/2017#,
      Nz([TermPenalty], 0),
      0
    )
  ) AS Adjustment,
  CCur(
    Nz([Rmkt].[BuyOutCost], 0)
  )+ CCur(
    IIf(
      [vw_Sch_Grp_Units_Remarketing].[ClientGroupId] = 432,
      Nz([TermPenalty], 0),
      0
    )
  ) AS [Final Cost],
  Rmkt.InvoicePaidDate,
  Rmkt.BuyOutDate,
  vw_Sch_Grp_Units_Remarketing.Status,
  Banks.BankShortName AS InterimBank,
  IIf(
    [IFPaidDate] <= [forms]![acctmenu]![AsOfDate],
    0,
    CCur(
      Nz([Rmkt.IFAmount], 0)
    )
  ) AS InterimAmt,
  Rmkt.IFLoanNumber AS AccountNumber,
  Rmkt.IFDate AS InterimDate,
  Rmkt.IFPaidDate AS InterimPaidDate,
  CCur(
    Nz([Final Cost], 0)- Nz([Rmkt].[IFAmount], 0)
  ) AS FAAmount,
  IIf(
    IsNull([Rmkt].[IFAmount]),
    Null,
    Round(
      1 -([Rmkt].[IFAmount] / [Final Cost]),
      3
    )
  ) AS FAPercent,
  Clients.ClientShNm AS Buyer,
  Rmkt.SoldPrice AS ActualSoldPrice,
  Units.unitoffleasedt AS ActualOffLease,
  Switch(
    [Wholesale] = 2, "OEM Trade", [Wholesale] = 0,
    "Retail", [Wholesale] = 1, "Wholesale"
  ) AS DSPN,
  vw_UnitInspectionLast.iPassedDate AS InspPassed
FROM
  (
    (
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
    )
    LEFT JOIN LookUp_UnitDesc ON Rmkt.UnitRef = LookUp_UnitDesc.AssetID
  )
  LEFT JOIN vw_UnitInspectionLast ON Units.UnitID = vw_UnitInspectionLast.AssetID
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Null
      Or (Rmkt.InvoicePaidDate)> forms!acctmenu!AsOfDate
    )
    And (
      (Rmkt.BuyOutDate)<= forms!acctmenu!AsOfDate
    )
  )
ORDER BY
  vw_Sch_Grp_Units_Remarketing.ClientShortName,
  vw_Sch_Grp_Units_Remarketing.AssetID,
  vw_Sch_Grp_Units_Remarketing.UnitNum;

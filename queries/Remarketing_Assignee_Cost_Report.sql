SELECT
  RemarketingReportAllUnits.SourceCompany,
  RemarketingReportAllUnits.Schedule,
  RemarketingReportAllUnits.UnitGroup AS [Group],
  RemarketingReportAllUnits.Unitnum AS [Unit#],
  RemarketingReportAllUnits.UnitVIN,
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.UnitYr,
  RemarketingReportAllUnits.UnitMake,
  RemarketingReportAllUnits.UnitModel,
  RemarketingReportAllUnits.UnitType,
  RemarketingReportAllUnits.Location,
  UnitOEC2.TotalCost AS [Original Cost],
  Units.UnitRent AS [Monthly Payment],
  SchGrp.BLCD AS BLCD,
  CInt([LeaseTermPrimary]) AS [Obligated Term],
  DateAdd(
    'm',
    CInt([LeaseTermPrimary]),
    [BLCD]
  )-1 AS [Expiration Date],
  RemarketingReportAllUnits.ParkDate AS [Parked Date],
  Units.unitoffleasedt AS [Off Lease Date],
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.BankPartner AS [Assignee Bank],
  RemarketingReportAllUnits.Odometer AS [Parked Odometer],
  Units.UnitSynResid AS Residual,
  RemarketingReportAllUnits.BuyOut AS [Buy Out Cost],
  RemarketingReportAllUnits.UnitSalePrice AS [Actual Sold Price],
  RemarketingReportAllUnits.TotalExpense AS [Remarketing Expenses],
  RemarketingReportAllUnits.TotalProfit AS [Gain/Loss On Sale],
  RemarketingReportAllUnits.ClientID,
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.InvoicePaid,
  RemarketingReportAllUnits.InvoiceSent,
  RemarketingReportAllUnits.BuyOutDate
FROM
  (
    (
      SchGrp
      INNER JOIN (
        Locations
        INNER JOIN (
          Units
          INNER JOIN RemarketingReportAllUnits ON Units.UnitID = RemarketingReportAllUnits.UnitID
        ) ON Locations.LocationID = RemarketingReportAllUnits.LocationID
      ) ON SchGrp.SGrpID = RemarketingReportAllUnits.SGrpID
    )
    INNER JOIN UnitOEC2 ON RemarketingReportAllUnits.UnitID = UnitOEC2.AssetId
  )
  LEFT JOIN RemarketReportExpenesTotal ON RemarketingReportAllUnits.UnitID = RemarketReportExpenesTotal.AssetId
WHERE
  (
    (
      (
        RemarketingReportAllUnits.SourceCompany
      )<> ""
    )
  );

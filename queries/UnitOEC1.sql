SELECT
  Units.UnitID AS AssetId,
  Units.UnitUnitNum AS UnitNo,
  Units.Unitvin,
  Units.UnitStatus,
  Units.GroupREF,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf(
    Nz([Axle], "TBD")= "TBD",
    Null,
    " " & [Axle]
  )& IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType] AS [Desc],
  Locations.Locationsn AS Location,
  Locations.LAddrLine1 AS StreetAddress,
  Locations.LCity AS City,
  Locations.LState AS State,
  Locations.LZip AS Zip,
  Nz(
    [Units].[UnitTitleNumber], [LastTitleLogEntry].[TitleNumber]
  ) AS [Title Num],
  Locations_1.LState AS TitledState,
  Units.unitstatus AS Status,
  Units.UnitInServ AS InService,
  CCur([UnitFinalCost]) AS UnitCost,
  Units.UnitInvoiceNum,
  Units.UnitVendorPytDate AS UnitInvoicePaid,
  Units.UnitIFBank AS InterimBank,
  Units.UnitIFDate AS InterimDate,
  Units.UnitIFAmt AS InterimAmt,
  Units.DrawDownId,
  Units.VendorREF,
  Units.UNITSRENT AS SYNDRent,
  Units.UNITPDRENT AS PerDiemRent,
  Units.UNITRENT AS LORent
FROM
  (
    (
      Units
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN Locations AS Locations_1 ON Units.UnitTitleLocationRef = Locations_1.LocationID
  )
  LEFT JOIN LastTitleLogEntry ON Units.UnitID = LastTitleLogEntry.AssetId
ORDER BY
  Units.Unitvin;

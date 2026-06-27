SELECT
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.CompanyName AS Lessee,
  vw_SixKeys.FAID,
  vw_SixKeys.UnitID AS Units,
  Nz(
    [LegalDescription], [UnitDescDocVer]
  ) AS UnitDesc,
  vw_SixKeys.UnitCost AS UnitCost,
  Nz(
    DSum(
      "ChildPartCost", "ChildPart_Cost_OEC_Unit",
      "UnitId=" & [Units].[UnitID]
    ),
    0
  ) AS [Child Cost],
  [UnitCost] + Nz(
    DSum(
      "ChildPartCost", "ChildPart_Cost_OEC_Unit",
      "UnitId=" & [Units].[UnitID]
    ),
    0
  ) AS OEC,
  Units.UnitVendorPytDate AS InvoicePaid,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SGprAllow1 AS [Allowed Miles],
  SchGrp.SGrpov1 AS MileChrg,
  SchGrp.SGprAllow2 AS MilesThreshold,
  SchGrp.SGrpov2 AS [MileChrg After],
  SchGrp.SGprSalesIndYr AS [Term Year],
  SchGrp.SgrpSalesIndRt AS Rate,
  SchGrp.SGrpSalesIndDt AS [Date],
  DLookUp(
    "StartsOn",
    "Lookup_LeaseStartsOn",
    "ID=" & Nz([SgrpCommInd], 10)
  ) AS [Neg BLCD],
  DLookUp(
    "StartsOn",
    "Lookup_LeaseStartsOn",
    "ID=" & Nz([PerDiemStartsOn], 10)
  ) AS [PerDiem Start],
  SchGrp.BLCD_Projected,
  vw_SixKeys.BankShortName AS Assignee,
  [SgrpSalesYield] - [SgrpSalesIndRt] AS Spread,
  SchGrp.SgrpSalesYield AS Yield,
  SchGrp.BPChgLO_Up AS Adjustment,
  vw_SixKeys.SchID,
  IIf(
    IsNull([Locations_1].[LAddrLine1]),
    Null,
    [Locations_1].[LAddrLine1] & " "
  )& IIf(
    IsNull([Locations_1].[LAddrLine2]),
    Null,
    [Locations_1].[LAddrLine2] & " "
  )& [Locations_1].[LCity] & " " & [Locations_1].[LState] & " " & [Locations_1].[LZip] & " (" & [Locations_1].[LCounty] & " County)" AS Domicle,
  IIf(
    IsNull([Locations].[LAddrLine1]),
    Null,
    [Locations].[LAddrLine1] & " "
  )& IIf(
    IsNull([Locations].[LAddrLine2]),
    Null,
    [Locations].[LAddrLine2] & " "
  )& [Locations].[LCity] & " " & [Locations].[LState] & " " & [Locations].[LZip] & " (" & [Locations].[LCounty] & " County)" AS DeliveryLocation,
  [StateSales] + [CountySales] + [CitySales] AS SalesTax,
  Units.UnitIRP AS IRP,
  vw_LocationTaxes.StateRReceipt AS RentReceipt,
  Nz(
    [StateException],
    Nz(
      [CountyException], [CityException]
    )
  ) AS AvalException,
  IIf(
    IsNull([Locations_1].[LAddrLine1]),
    Null,
    [Locations_2].[LAddrLine1] & " "
  )& IIf(
    IsNull([Locations_2].[LAddrLine2]),
    Null,
    [Locations_2].[LAddrLine2] & " "
  )& [Locations_2].[LCity] & " " & [Locations_2].[LState] & " " & [Locations_2].[LZip] AS Billing,
  SchGrp.BillingInstructions,
  SchGrp.TaxInstructions,
  IIf(
    IsNull([Locations_1].[LAddrLine1]),
    Null,
    [Locations_3].[LAddrLine1] & " "
  )& IIf(
    IsNull([Locations_3].[LAddrLine2]),
    Null,
    [Locations_3].[LAddrLine2] & " "
  )& [Locations_3].[LCity] & " " & [Locations_3].[LState] & " " & [Locations_3].[LZip] AS Title,
  Units.UnitSRent AS SchRent,
  Units.FirstExtRent AS ExtdRent,
  Units.AssetID_FATitleLog AS FATitleLogNum,
  vw_SixKeys.VendorCost,
  vw_SixKeys.VIN,
  vw_SixKeys.UnitNum
FROM
  (
    (
      (
        (
          (
            (
              vw_SixKeys
              INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
            )
            INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
          )
          LEFT JOIN Locations ON Units.UnitDeliveryLocationRef = Locations.LocationID
        )
        LEFT JOIN Locations AS Locations_1 ON Units.UnitGarageLocationRef = Locations_1.LocationID
      )
      LEFT JOIN vw_LocationTaxes ON Locations_1.LocationID = vw_LocationTaxes.LocationID
    )
    LEFT JOIN Locations AS Locations_2 ON SchGrp.BillToAddressID = Locations_2.LocationID
  )
  LEFT JOIN Locations AS Locations_3 ON Units.UnitTitleLocationRef = Locations_3.LocationID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  );

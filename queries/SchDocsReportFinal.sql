SELECT
  SU.[Lease No],
  SU.Lessee,
  SU.Co_Lessee,
  SU.SchNo,
  SU.Group,
  SU.[Unit No],
  SU.Unitvin,
  SU.Desc,
  SU.UnitCost,
  SU.VendorCost,
  SU.UnitInvoiceNum AS [Invoice Num],
  SU.UnitInvoiceDate AS [Invoice Date],
  SU.UnitInvoicePaid AS [Invoice Paid],
  Cost_ChildParts_Pivot_Make.[Dealer Mod
ECU Programming Make],
  Cost_ChildParts_Pivot_Cost.[Dealer Mod
ECU Programming Cost],
  Cost_ChildParts_Pivot_InvoiceNum.[Dealer Mod
ECU Programming Invoice],
  Cost_ChildParts_Pivot_InvoiceDate.[Dealer Mod
ECU Programming Paid],
  Cost_ChildParts_Pivot_DrawNum.[Dealer Mod
ECU Programming DrawNum],
  Cost_ChildParts_Pivot_Make.[Doc Fee
Plate Make],
  Cost_ChildParts_Pivot_Cost.[Doc Fee
Plate Cost],
  Cost_ChildParts_Pivot_InvoiceNum.[Doc Fee
Plate Invoice],
  Cost_ChildParts_Pivot_InvoiceDate.[Doc Fee
Plate Paid],
  Cost_ChildParts_Pivot_DrawNum.[Doc Fee
Plate DrawNum],
  Cost_ChildParts_Pivot_Make.[Doc Fee 2
Title /Reg Make],
  Cost_ChildParts_Pivot_Cost.[Doc Fee 2
Title /Reg Cost],
  Cost_ChildParts_Pivot_InvoiceNum.[Doc Fee 2
Title /Reg Invoice],
  Cost_ChildParts_Pivot_InvoiceDate.[Doc Fee 2
Title /Reg Paid],
  Cost_ChildParts_Pivot_DrawNum.[Doc Fee 2
Title /Reg DrawNum],
  Cdbl(
    nz(
      [Dealer Mod
ECU Programming Cost],
      0
    )
  )+ Cdbl(
    nz([Doc Fee
Plate Cost], 0)
  )+ Cdbl(
    nz([Doc Fee 2
Title /Reg Cost], 0)
  ) AS ChildPartTotal,
  [UnitCost] + [ChildPartTotal] AS [Unit OEC],
  SU.AssetId,
  SU.Status,
  SU.FAID,
  SU.EstiDeliveryDate AS [Est Delv],
  SU.ActualDelDate AS [Actual Delv],
  SU.DeliveryLoc,
  SU.InService,
  SU.[Title Num],
  SU.TitledState,
  SU.Location AS Domicile,
  SU.StreetAddress AS Address,
  SU.City,
  SU.State,
  SU.Zip,
  SU.County,
  SU.InterimBank,
  SU.InterimDate AS [Interim Funded],
  SU.InterimAmt,
  SU.DrawName,
  SU.CustodianID,
  SU.TaxExempt AS [Tax Exempt],
  SU.UpFront AS [Up Front],
  SU.UnitTaxNotes,
  SU.[State Tax],
  SU.[County Tax],
  SU.[City Tax],
  SU.[Other Tax],
  SU.[Loc Tax] AS [Total Tax],
  SU.LTaxRate AS [ATLAAS TaxRate]
FROM
  ScheduleUnitDocs3 AS SU
  LEFT JOIN Cost_ChildParts_AllTypes ON SU.AssetId = Cost_ChildParts_AllTypes.AssetId
WHERE
  (
    (
      (SU.AssetID) Is Not Null
    )
  )
ORDER BY
  SU.Group,
  SU.[Unit No],
  SU.AssetId;

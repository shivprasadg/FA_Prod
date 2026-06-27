SELECT
  Cost_ChildParts_AssetID.AssetId,
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
  ) AS ChildPartTotal
FROM
  (
    (
      (
        (
          Cost_ChildParts_AssetID
          LEFT JOIN Cost_ChildParts_Pivot_Make ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_Make.AssetID
        )
        LEFT JOIN Cost_ChildParts_Pivot_Cost ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_Cost.AssetID
      )
      LEFT JOIN Cost_ChildParts_Pivot_InvoiceNum ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_InvoiceNum.AssetID
    )
    LEFT JOIN Cost_ChildParts_Pivot_InvoiceDate ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_InvoiceDate.AssetID
  )
  LEFT JOIN Cost_ChildParts_Pivot_DrawNum ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_DrawNum.AssetID;

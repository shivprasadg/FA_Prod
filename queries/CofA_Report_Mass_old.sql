SELECT
  CofA_Report_Mass_Pre.*,
  CofA_ChildParts.ChildYear,
  IIf(
    [ChildMake] <> "", [ChildMake] & " ",
    ""
  )& IIf(
    [ChildClass] <> "", [ChildClass] & " ",
    ""
  ) AS ChildMakeModel,
  CofA_ChildParts.ChildVIN,
  CofA_ChildParts.childtype,
  IIf(
    [childtype] = "Refrig Unit", "Refrigeration Unit",
    [ChildType]
  ) AS Ctype,
  [clientCompanyName] & Chr(13)& Chr(10)& IIf(
    IsNull([clientCoNm2]),
    "",
    [clientCoNm2] & Chr(13)& Chr(10)
  )& [clientAddress1] & " " & [clientAddress2] & Chr(13)& Chr(10)& [clientcity] & ", " & [clientstate] & " " & [clientzip] AS Lessee,
  UnitTotalOEC.SalesRent AS Rentx,
  UnitTotalOEC.UnitTotalCost AS TotalCost,
  CofA_ChildParts.ChildPartID,
  [clientCompanyName] & Chr(13)& Chr(10)& [clientAddress1] & " " & [clientAddress2] & Chr(13)& Chr(10)& [clientcity] & ", " & [clientstate] & " " & [clientzip] AS Lesseez
FROM
  (
    CofA_Report_Mass_Pre
    LEFT JOIN CofA_ChildParts ON CofA_Report_Mass_Pre.UnitID = CofA_ChildParts.AssetId
  )
  LEFT JOIN UnitTotalOEC ON CofA_Report_Mass_Pre.UnitID = UnitTotalOEC.AssetID;

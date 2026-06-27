SELECT
  vw_OPS_Report_Base.Lessee,
  vw_OPS_Report_Base.[Lease No],
  vw_OPS_Report_Base.SchNo,
  vw_OPS_Report_Base.[Group],
  vw_OPS_Report_Base.[Unit No],
  vw_OPS_Report_Base.AssetId,
  vw_OPS_Report_Base.Status,
  vw_OPS_Report_Base.UnitVIN,
  vw_OPS_Report_Base.[Desc],
  vw_OPS_Report_Base.Garage,
  vw_OPS_Report_Base.StreetAddress,
  vw_OPS_Report_Base.City,
  vw_OPS_Report_Base.State,
  vw_OPS_Report_Base.Zip,
  vw_OPS_Report_Base.County,
  vw_OPS_Report_Base.RegLoc,
  cRefrig.[Refrig Make],
  cRefrig.[Refrig Class Type],
  cRefrig.[Refrig Serial],
  vw_OPS_Report_Base.PlateNum,
  vw_OPS_Report_Base.Delivered,
  vw_OPS_Report_Base.Accepted,
  vw_OPS_Report_Base.InServiceInd,
  vw_OPS_Report_Base.InServiceDate,
  vw_OPS_Report_Base.MSORecv,
  vw_OPS_Report_Base.MSOCopy,
  vw_OPS_Report_Base.TitleAppSent,
  vw_OPS_Report_Base.LastTitleLog,
  vw_OPS_Report_Base.TitleAction,
  vw_OPS_Report_Base.UnitCost,
  vw_OPS_Report_Base.ChildPartCost,
  vw_OPS_Report_Base.UnitTotalCost,
  vw_OPS_Report_Base.VendorCost,
  vw_OPS_Report_Base.Residual,
  vw_OPS_Report_Base.Rent,
  vw_OPS_Report_Base.SyndRent,
  vw_OPS_Report_Base.PerDiemStart,
  vw_OPS_Report_Base.BLCD,
  vw_OPS_Report_Base.ProjEOL,
  vw_OPS_Report_Base.ODORead,
  vw_OPS_Report_Base.ODOSource,
  vw_OPS_Report_Base.DeliveryLoc,
  vw_OPS_Report_Base.Title,
  vw_OPS_Report_Base.TitleState,
  vw_OPS_Report_Base.LienRelease,
  vw_OPS_Report_Base.CustodianID,
  vw_OPS_Report_Base.SchID
FROM
  (
    vw_OPS_Report_Base
    LEFT JOIN Ops_ChildPart_Make ON Ops_ChildPart_Make.AssetID = vw_OPS_Report_Base.AssetId
  )
  LEFT JOIN (
    SELECT
      UnitRef AS cAssetID, ChildMake AS [Refrig Make],
      ChildClass & ' ' & ChildType AS [Refrig Class Type],
      ChildVIN AS [Refrig Serial]
    FROM
      ChildParts
    WHERE
      ChildClass LIKE '*Temp*'
      AND ChildClass LIKE '*Refrig*'
  ) AS cRefrig ON cRefrig.cAssetID = vw_OPS_Report_Base.AssetId;

SELECT
  CustTitleImport.[Record #],
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  CustTitleImport.[Alternate ID],
  CustTitleImport.[Loan Amount],
  CustTitleImport.[Trust Number],
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Status
FROM
  vw_SixKeys
  INNER JOIN CustTitleImport ON vw_SixKeys.VIN = CustTitleImport.VIN
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Status;

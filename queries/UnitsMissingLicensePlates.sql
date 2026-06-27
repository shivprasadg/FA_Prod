SELECT
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLNo,
  vw_SixKeys.BankShortName AS Assignee,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup AS [Group],
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum,
  vw_SixKeys.VIN,
  Units.UnitPlateNum,
  vw_SixKeys.AcceptedDate,
  Units.UnitCOASent,
  Nz([UnitCOABack], [UnitCofaCopy]) AS [COAOrig/Copy],
  vw_SixKeys.Status,
  vw_SixKeys.UnitDocAccpt,
  LookUp_UnitDesc.DescriptionMed AS [Desc],
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  (
    vw_SixKeys
    INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
  )
  INNER JOIN LookUp_UnitDesc ON Units.UnitID = LookUp_UnitDesc.AssetID
WHERE
  (
    (
      (Units.UnitPlateNum) Is Null
    )
    AND (
      (vw_SixKeys.Status)= "A"
      Or (vw_SixKeys.Status)= "C"
    )
  )
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.AssetID;

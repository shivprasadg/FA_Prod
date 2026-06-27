SELECT
  DrawDowns.DrawDownId,
  Banks.BankShortName,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawNumber,
  DrawDowns.DrawName,
  Count(vw_SixKeys.AssetID) AS Units,
  Count(ChildParts.ChildPartID) AS Parts,
  Sum(vw_SixKeys.UnitCost) AS UnitAmt,
  Sum(ChildParts.childFinalCost) AS PartAmt
FROM
  (
    (
      DrawDowns
      LEFT JOIN vw_SixKeys ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
    )
    LEFT JOIN ChildParts ON DrawDowns.DrawDownId = ChildParts.DrawDownId
  )
  INNER JOIN Banks ON DrawDowns.DrawBankId = Banks.BankId
WHERE
  (
    (
      (DrawDowns.MLID)<> 0
    )
  )
GROUP BY
  DrawDowns.DrawDownId,
  Banks.BankShortName,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawNumber,
  DrawDowns.DrawName
ORDER BY
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawNumber;

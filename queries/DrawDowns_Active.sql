SELECT
  Nz([banks].[BankShortName], "TBD") AS Bank,
  DrawDowns.DrawDownId,
  DrawDownType.DrawnDownType AS Type,
  MstrLease.MLNo,
  DrawDowns.DrawRequestNumber AS Request,
  DrawDowns.DrawNumber AS Draw,
  Count(Units.Unitid) AS Units,
  Count(ChildParts.ChildPartId) AS Child,
  MstrLease.MLID,
  Banks.BankId,
  DrawDowns.ClientID,
  DrawDowns.DrawNumber,
  DrawDowns.DateDrawn
FROM
  (
    (
      (
        (
          DrawDowns
          LEFT JOIN Units ON DrawDowns.DrawDownId = Units.DrawDownId
        )
        LEFT JOIN ChildParts ON DrawDowns.DrawDownId = ChildParts.DrawDownId
      )
      LEFT JOIN Banks ON DrawDowns.DrawBankId = Banks.bankID
    )
    INNER JOIN MstrLease ON DrawDowns.MLID = MstrLease.MLID
  )
  INNER JOIN DrawDownType ON DrawDowns.DrawTypeId = DrawDownType.DrawDownTypeId
GROUP BY
  Nz([banks].[BankShortName], "TBD"),
  DrawDowns.DrawDownId,
  DrawDownType.DrawnDownType,
  MstrLease.MLNo,
  DrawDowns.DrawRequestNumber,
  MstrLease.MLID,
  Banks.BankId,
  DrawDowns.ClientID,
  DrawDowns.DrawNumber,
  DrawDowns.DateDrawn
ORDER BY
  Nz([banks].[BankShortName], "TBD"),
  DrawDowns.DrawDownId,
  DrawDowns.DrawRequestNumber;

SELECT
  Banks.BankId,
  IIf(
    [mlorig] <> "FA"
    And IsNull([banks].[BankShortName]),
    "Not FA",
    [banks].[BankShortName]
  ) AS Assignee,
  Count(vw_SixKeys.AssetID) AS Units,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  (
    SchGrp
    INNER JOIN vw_SixKeys ON SchGrp.SGrpID = vw_SixKeys.SGrpID
  )
  LEFT JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
GROUP BY
  Banks.BankId,
  IIf(
    [mlorig] <> "FA"
    And IsNull([banks].[BankShortName]),
    "Not FA",
    [banks].[BankShortName]
  ),
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID;

SELECT
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.AssetID,
  vw_SixKeys.Status,
  ExitOptionsGroup_EOL.EOLDate
FROM
  vw_SixKeys
  LEFT JOIN ExitOptionsGroup_EOL ON vw_SixKeys.AssetID = ExitOptionsGroup_EOL.AssetID
WHERE
  (
    (
      (vw_SixKeys.Status)= "TR"
    )
  )
ORDER BY
  ExitOptionsGroup_EOL.EOLDate DESC;

SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.CompanyName,
  Count(vw_SixKeys.AssetID) AS [Unit Count]
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.Status)= "A"
      Or (vw_SixKeys.Status)= "B"
      Or (vw_SixKeys.Status)= "O"
      Or (vw_SixKeys.Status)= "C"
      Or (vw_SixKeys.Status)= "TR"
    )
  )
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.CompanyName
ORDER BY
  vw_SixKeys.CompanyName;

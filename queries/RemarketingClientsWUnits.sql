SELECT
  clientID,
  ClientShortName,
  CompanyName,
  Count(AssetID) AS [Unit Count]
FROM
  vw_SixKeys
GROUP BY
  clientID,
  ClientShortName,
  CompanyName
ORDER BY
  CompanyName;

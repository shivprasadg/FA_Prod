SELECT
  vw_SixKeys.CompanyName,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.FAID,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS Units
FROM
  vw_SixKeys
GROUP BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.FAID,
  vw_SixKeys.MLOrig
ORDER BY
  vw_SixKeys.CompanyName;

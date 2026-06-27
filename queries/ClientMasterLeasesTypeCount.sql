SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLNo,
  IIf([MLOrig] = "FA", 1, 0) AS FA,
  IIf([MLOrig] <> "FA", 1, 0) AS NotFA,
  Count(vw_SixKeys.UnitId) AS Units,
  vw_SixKeys.ClientGroupId
FROM
  vw_SixKeys
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLNo,
  vw_SixKeys.ClientGroupId
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig;

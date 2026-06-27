SELECT
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.CompanyType,
  vw_SixKeys.FAID,
  vw_SixKeys.MLNo,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Schedule,
  Count(vw_SixKeys.AssetID) AS Units,
  CDbl(
    20 & Trim(
      Left([FAID], 2)
    )
  ) AS [Year]
FROM
  vw_SixKeys
GROUP BY
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.CompanyType,
  vw_SixKeys.FAID,
  vw_SixKeys.MLNo,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Schedule
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID;

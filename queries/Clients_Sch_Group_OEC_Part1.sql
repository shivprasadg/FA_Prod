SELECT
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLNo,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  Count(vw_SixKeys.AssetID) AS Units
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLNo,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.clientID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID;

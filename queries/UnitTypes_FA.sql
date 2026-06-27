SELECT
  vw_SixKeys.MLOrig,
  vw_SixKeys.Type,
  Count(vw_SixKeys.AssetID) AS CountOfAssetID
FROM
  vw_SixKeys
GROUP BY
  vw_SixKeys.MLOrig,
  vw_SixKeys.Type
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  );

SELECT
  vw_SixKeys.ClientShortName,
  Count(vw_SixKeys.AssetID) AS UnitsInServiced,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Min(vw_SixKeys.InServiceDate) AS FirstInServiced,
  Max(vw_SixKeys.InServiceDate) AS LastInServiced
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.InServiceDate)>= '1/1/' & [AskYear]
      And (vw_SixKeys.InServiceDate)< '12/31/' & [AskYear]
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId;

SELECT
  vw_SixKeys.ClientShortName,
  Count(vw_SixKeys.AssetID) AS UnitsOffLease,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Min(vw_SixKeys.OffLeaseDate) AS FirstOL,
  Max(vw_SixKeys.OffLeaseDate) AS LastOL
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.OffLeaseDate)>= '1/1/' & [AskYear]
      And (vw_SixKeys.OffLeaseDate)< '12/31/' & [AskYear]
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId;

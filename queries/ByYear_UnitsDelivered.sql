SELECT
  vw_SixKeys.ClientShortName,
  Count(vw_SixKeys.AssetID) AS UnitDelivered,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Min(vw_SixKeys.ActualDeliveryDate) AS FirstDelivered,
  Max(vw_SixKeys.ActualDeliveryDate) AS LastDelivered
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.ActualDeliveryDate)>= '1/1/' & [AskYear]
      And (vw_SixKeys.ActualDeliveryDate)< '12/31/' & [AskYear]
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId;

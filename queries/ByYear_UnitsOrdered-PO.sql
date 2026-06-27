SELECT
  vw_SixKeys.ClientShortName,
  Count(vw_SixKeys.AssetID) AS UnitOrdered,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Min(vw_SixKeys.PurchaseOrderDate) AS FirstPO,
  Max(vw_SixKeys.PurchaseOrderDate) AS LastPO
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.PurchaseOrderDate)>= '1/1/' & [AskYear]
      And (vw_SixKeys.PurchaseOrderDate)< '12/31/' & [AskYear]
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId;

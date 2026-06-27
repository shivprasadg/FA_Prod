SELECT
  AssetID
FROM
  Cost_ChildParts
WHERE
  (
    (
      (Cost_ChildParts.SchID)= 3056
    )
  )
GROUP BY
  Cost_ChildParts.AssetId;

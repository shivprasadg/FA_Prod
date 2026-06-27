TRANSFORM First(Cost_ChildParts.PartCost) AS FirstOfPartCost
SELECT
  Cost_ChildParts.AssetID
FROM
  Cost_ChildParts
WHERE
  (
    (
      (Cost_ChildParts.SchID)= 3015
    )
  )
GROUP BY
  Cost_ChildParts.AssetID PIVOT [PartMake] & Chr(13)& Chr(10)& [PartType] & ' Cost';

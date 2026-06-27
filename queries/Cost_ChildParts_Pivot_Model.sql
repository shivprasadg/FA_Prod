TRANSFORM First(Cost_ChildParts.PartType) AS FirstOfPartModel
SELECT
  Cost_ChildParts.AssetID
FROM
  Cost_ChildParts
GROUP BY
  Cost_ChildParts.AssetID PIVOT [PartClass] & " Type";

TRANSFORM First(Cost_ChildParts.PartInvoice) AS FirstOfPartInvoice
SELECT
  Cost_ChildParts.AssetID
FROM
  Cost_ChildParts
WHERE
  (
    (
      (Cost_ChildParts.SchID)= 3056
    )
  )
GROUP BY
  Cost_ChildParts.AssetID PIVOT [PartClass] & Chr (13)& Chr (10)& [PartType] & ' Invoice';

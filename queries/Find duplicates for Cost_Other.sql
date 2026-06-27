SELECT
  Cost_Other.AssetId,
  Cost_Other.OtherMake,
  Cost_Other.ChildType,
  Cost_Other.OtherInvoice
FROM
  Cost_Other
WHERE
  (
    (
      (Cost_Other.AssetId) In (
        SELECT
          [AssetId]
        FROM
          [Cost_Other] As Tmp
        GROUP BY
          [AssetId]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  Cost_Other.AssetId;

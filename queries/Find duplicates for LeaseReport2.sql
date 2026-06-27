SELECT
  LeaseReport2.AssetId,
  LeaseReport2.FAID,
  LeaseReport2.UnitType,
  LeaseReport2.Make
FROM
  LeaseReport2
WHERE
  (
    (
      (LeaseReport2.AssetId) In (
        SELECT
          [AssetId]
        FROM
          [LeaseReport2] As Tmp
        GROUP BY
          [AssetId]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  LeaseReport2.AssetId;

SELECT
  TaxStatesBase.[TaxTypeId],
  TaxStatesBase.[TaxStatesID],
  TaxStatesBase.[StartsOn],
  TaxStatesBase.[EndsOn]
FROM
  TaxStatesBase
WHERE
  (
    (
      (TaxStatesBase.[TaxTypeId]) In (
        SELECT
          [TaxTypeId]
        FROM
          [TaxStatesBase] As Tmp
        GROUP BY
          [TaxTypeId]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  TaxStatesBase.[TaxTypeId];

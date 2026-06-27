SELECT
  Rmkt.[UnitRef],
  Rmkt.[SUF Sent],
  Rmkt.[SUF Returned],
  Rmkt.[BuyOutDate],
  Rmkt.[BuyOutCost],
  Rmkt.[BuyerRef]
FROM
  Rmkt
WHERE
  (
    (
      (Rmkt.[UnitRef]) In (
        SELECT
          [UnitRef]
        FROM
          [Rmkt] As Tmp
        GROUP BY
          [UnitRef]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  Rmkt.[UnitRef];

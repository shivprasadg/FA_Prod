SELECT
  Lines.Lineid,
  Sum(PROMNT.PNAmt) AS SumOfPNAmt
FROM
  Lines
  INNER JOIN PROMNT ON Lines.Lineid = PROMNT.LineID
WHERE
  (
    (
      (PROMNT.PNExp)> Date()
    )
  )
GROUP BY
  Lines.Lineid
ORDER BY
  Lines.Lineid;

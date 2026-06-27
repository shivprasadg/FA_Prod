SELECT
  ExitOptionSelected_A.UnitID,
  Max(
    Nz([Exch], "----")
  ) AS Exc,
  Max(
    Nz([Extend], "----")
  ) AS Extd,
  Max(
    Nz([Term], "----")
  ) AS [End]
FROM
  ExitOptionSelected_A
GROUP BY
  ExitOptionSelected_A.UnitID;

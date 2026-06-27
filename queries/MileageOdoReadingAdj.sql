SELECT
  vw_UMP_Mileage.UnitID,
  Min(vw_UMP_Mileage.LogActivityDate) AS FirstReport,
  Max(vw_UMP_Mileage.LogActivityDate) AS LastReport,
  DateDiff(
    'm',
    Min([LogActivityDate]),
    Max([LogActivityDate])
  ) AS MthsRptd,
  Round(
    Max([OdometerEnd])- Min([OdometerBegin]),
    1
  ) AS ODOMiles,
  IIf(
    Nz([ODOMiles], 0)> 0
    And Nz([MthsRptd], 0)> 0,
    Round([ODOMiles] / [MthsRptd], 1),
    0
  ) AS MPM
FROM
  vw_UMP_Mileage
WHERE
  (
    (
      (vw_UMP_Mileage.OBCDataStatus)<> 6
    )
  )
GROUP BY
  vw_UMP_Mileage.UnitID;

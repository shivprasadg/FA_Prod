SELECT
  Units.UnitGarageLocationRef,
  Min(
    IIf(
      [UnitStatus] = "B",
      Nz(
        [UnitActualDelDate],
        Nz([UnitEstDelDate])
      ),
      ''
    )
  ) AS FirstDelvery,
  Max(
    IIf(
      [UnitStatus] = "B",
      Nz(
        [UnitActualDelDate],
        Nz([UnitEstDelDate])
      ),
      ''
    )
  ) AS LastDelvery
FROM
  Units
WHERE
  (
    (
      (Units.Unitstatus)= "B"
    )
  )
GROUP BY
  Units.UnitGarageLocationRef;

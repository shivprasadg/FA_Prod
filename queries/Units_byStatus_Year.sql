SELECT
  DatePart("yyyy", [UnitBuildDate]) AS BuildYear,
  Units.UnitStatus,
  IIf(
    [UnitStatus] = "A",
    Count([UnitStatus]),
    0
  ) AS Active,
  IIf(
    [UnitStatus] = "B",
    Count([UnitStatus]),
    0
  ) AS BeingBuilt,
  IIf(
    [UnitStatus] = "C",
    Count([UnitStatus]),
    0
  ) AS Casualty,
  IIf(
    [UnitStatus] = "K",
    Count([UnitStatus]),
    0
  ) AS Brokered,
  IIf(
    [UnitStatus] = "KS",
    Count([UnitStatus]),
    0
  ) AS SoldNotOwn,
  IIf(
    [UnitStatus] = "O",
    Count([UnitStatus]),
    0
  ) AS OffLease,
  IIf(
    [UnitStatus] = "R",
    Count([UnitStatus]),
    0
  ) AS Remarketing,
  IIf(
    [UnitStatus] = "S",
    Count([UnitStatus]),
    0
  ) AS SoldOwn,
  IIf(
    [UnitStatus] = "T",
    Count([UnitStatus]),
    0
  ) AS Tracking,
  IIf(
    [UnitStatus] = "TR",
    Count([UnitStatus]),
    0
  ) AS TrackPending
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (vw_SixKeys.mlorig)= "FA"
    )
  )
GROUP BY
  DatePart("yyyy", [UnitBuildDate]),
  Units.UnitStatus
ORDER BY
  DatePart("yyyy", [UnitBuildDate]);

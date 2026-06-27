SELECT
  LeaseReport2.Lessee,
  LeaseReport2.[Lease No],
  IIf(
    [UnitType] Like "*trailer*", "Trailer",
    "Tractor"
  ) AS UnitClass,
  LeaseReport2.ModelYear,
  LeaseReport2.Make,
  LeaseReport2.UnitType,
  CDbl(
    Nz([LeaseReport2].[OrderedYear], 0)
  ) AS OrderedYear,
  CDbl(
    Nz([LeaseReport2].[OwnedYear], 0)
  ) AS OwnedYear,
  CDbl(
    Nz([LeaseReport2].[SYDYear], 0)
  ) AS SYDYear,
  LeaseReport2.Ordered,
  IIf(
    [Ordered] = "Ordered",
    Count([AssetId]),
    0
  ) AS OnOrderCount,
  IIf(
    [Ordered] = "Ordered",
    Sum(
      CCur(
        Nz([LeaseReport2].[UnitCost], 0)
      )
    ),
    0
  ) AS OnOrderCost,
  IIf(
    [Ordered] = "Ordered",
    Sum(
      CCur(
        Nz([LeaseReport2].[TotalCost], 0)
      )
    ),
    0
  ) AS OnOrderOEC,
  LeaseReport2.Owned,
  IIf(
    [Owned] = "Owned",
    Count([AssetId]),
    0
  ) AS OwnedCount,
  IIf(
    [Owned] = "Owned",
    Sum(
      CCur(
        Nz([LeaseReport2].[UnitCost], 0)
      )
    ),
    0
  ) AS OwnedCost,
  IIf(
    [Owned] = "Owned",
    Sum(
      CCur(
        Nz([LeaseReport2].[TotalCost], 0)
      )
    ),
    0
  ) AS OwnedOEC,
  LeaseReport2.Syndicated,
  IIf(
    [Syndicated] = "Syndicated",
    Count([AssetId]),
    0
  ) AS SydCount,
  IIf(
    [Syndicated] = "Syndicated",
    Sum(
      CCur(
        Nz([LeaseReport2].[UnitCost], 0)
      )
    ),
    0
  ) AS SydCost,
  IIf(
    [Syndicated] = "Syndicated",
    Sum(
      CCur(
        Nz([LeaseReport2].[TotalCost], 0)
      )
    ),
    0
  ) AS SydOEC
FROM
  LeaseReport2
GROUP BY
  LeaseReport2.Lessee,
  LeaseReport2.[Lease No],
  IIf(
    [UnitType] Like "*trailer*", "Trailer",
    "Tractor"
  ),
  LeaseReport2.ModelYear,
  LeaseReport2.Make,
  LeaseReport2.UnitType,
  CDbl(
    Nz([LeaseReport2].[OrderedYear], 0)
  ),
  CDbl(
    Nz([LeaseReport2].[OwnedYear], 0)
  ),
  CDbl(
    Nz([LeaseReport2].[SYDYear], 0)
  ),
  LeaseReport2.Ordered,
  LeaseReport2.Owned,
  LeaseReport2.Syndicated
ORDER BY
  IIf(
    [UnitType] Like "*trailer*", "Trailer",
    "Tractor"
  ),
  CDbl(
    Nz([LeaseReport2].[OrderedYear], 0)
  ) DESC,
  CDbl(
    Nz([LeaseReport2].[OwnedYear], 0)
  ) DESC,
  CDbl(
    Nz([LeaseReport2].[SYDYear], 0)
  ) DESC;

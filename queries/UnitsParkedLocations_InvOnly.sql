SELECT
  vw_SixKeys.LocationId AS GarageLocationID,
  vw_SixKeys.AssetID,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitNum,
  vw_SixKeys.UnitDescShortVer AS [Desc],
  vw_SixKeys.MYear,
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  vw_SixKeys.Type,
  vw_SixKeys.SubType,
  vw_SixKeys.Status,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.ClientID,
  vw_SixKeys.MLOrig,
  IIf(
    [status] = "A"
    Or [status] = "B"
    Or [status] = "R"
    Or [status] = "T"
    Or [status] = "TR",
    1, 0
  ) AS InFleet
FROM
  vw_SixKeys
WHERE
  (
    (
      (
        IIf(
          [status] = "A"
          Or [status] = "B"
          Or [status] = "R"
          Or [status] = "T"
          Or [status] = "TR",
          1, 0
        )
      )= 1
    )
  )
ORDER BY
  vw_SixKeys.MLOrig,
  vw_SixKeys.UnitID;

SELECT
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  IIf(
    [vw_SixKeys].[MLOrig] = "FA",
    Count([UnitID]),
    0
  ) AS Leased,
  IIf(
    [vw_SixKeys].[MLOrig] = "FA",
    0,
    Count([UnitID])
  ) AS Tracked,
  vw_SixKeys.MLOrig
FROM
  vw_SixKeys
GROUP BY
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  vw_SixKeys.MLOrig
ORDER BY
  vw_SixKeys.Make,
  vw_SixKeys.Model;

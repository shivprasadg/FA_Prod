SELECT
  vw_SixKeys.MLOrig,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLID,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.UnitDescShortVer,
  Count(vw_SixKeys.assetID) AS Units,
  Units.UnitsRent AS SchRent,
  Units.UnitRent AS SyndRent,
  IIf(
    [Units].[UnitRent] < [Units].[UnitSRent],
    ">SchRent", ''
  ) AS SchVsSynd
FROM
  vw_SixKeys
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
GROUP BY
  vw_SixKeys.MLOrig,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLID,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.UnitDescShortVer,
  Units.UnitsRent,
  Units.UnitRent,
  IIf(
    [Units].[UnitRent] < [Units].[UnitSRent],
    ">SchRent", ''
  ),
  vw_SixKeys.SchID
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (Units.UnitsRent)> 0
    )
    AND (
      (Units.UnitRent)> 0
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.SchID;

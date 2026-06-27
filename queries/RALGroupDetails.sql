SELECT
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  Count(vw_SixKeys.UnitID) AS Units,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Sum(UnitTotalOEC.UnitTotalCost) AS Lessorcost,
  UnitDescinGroup([vw_sixkeys].[sGrpid], True) AS [Desc],
  UnitRentsinGroup(vw_SixKeys.[SGrpID], True) AS SyndRent,
  Format(
    [vw_sixkeys].[BLCD], "mmmm dd"", ""yyyy"
  ) AS BLCD,
  vw_SixKeys.LeaseTermPrimary
FROM
  RALTemp1
  INNER JOIN (
    vw_SixKeys
    INNER JOIN UnitTotalOEC ON vw_SixKeys.AssetID = UnitTotalOEC.AssetID
  ) ON RALTemp1.SchID = vw_SixKeys.SchID
GROUP BY
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Format(
    [vw_sixkeys].[BLCD], "mmmm dd"", ""yyyy"
  ),
  vw_SixKeys.LeaseTermPrimary;

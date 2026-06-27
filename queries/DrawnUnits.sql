SELECT
  vw_SixKeys.FAID,
  vw_SixKeys.SchId,
  Count(vw_SixKeys.AssetId) AS DrawnUnits
FROM
  DrawDowns
  INNER JOIN vw_SixKeys ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
GROUP BY
  vw_SixKeys.FAID,
  vw_SixKeys.SchId;

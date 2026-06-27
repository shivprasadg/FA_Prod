SELECT
  DDU.DrawDownId,
  DDU.RequestNumber AS [Request#],
  DDU.Draw,
  DDU.DrawName,
  Count(DDU.AssetID) AS Units,
  Sum(DDU.UnitCost) AS UnitAmt,
  DDU.SchID
FROM
  DrawDowns_Units_All AS DDU
WHERE
  (
    (
      Not (DDU.DrawDownId) Is Null
    )
  )
GROUP BY
  DDU.DrawDownId,
  DDU.RequestNumber,
  DDU.Draw,
  DDU.DrawName,
  DDU.SchID
ORDER BY
  DDU.Draw;

SELECT
  DrawDowns_Units_Drawn.DrawDownId,
  DrawDowns_Units_Drawn.RequestNumber AS [Request#],
  DrawDowns_Units_Drawn.Draw,
  DrawDowns_Units_Drawn.DrawName,
  Count(DrawDowns_Units_Drawn.AssetID) AS Units,
  Format(
    Sum([UnitCost]),
    "$#,###"
  ) AS UnitAmt,
  DrawDowns_Units_Drawn.Invoice,
  DrawDowns_Units_Drawn.SGrpID
FROM
  DrawDowns_Units_Drawn
GROUP BY
  DrawDowns_Units_Drawn.DrawDownId,
  DrawDowns_Units_Drawn.RequestNumber,
  DrawDowns_Units_Drawn.Draw,
  DrawDowns_Units_Drawn.DrawName,
  DrawDowns_Units_Drawn.Invoice,
  DrawDowns_Units_Drawn.SGrpID
ORDER BY
  DrawDowns_Units_Drawn.Draw;

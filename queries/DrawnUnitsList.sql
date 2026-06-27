SELECT
  DISTINCTROW vw_SixKeys.AssetId,
  vw_SixKeys.VIN,
  vw_SixKeys.Unitnum,
  vw_SixKeys.Schedule AS SchNum,
  vw_SixKeys.InvoiceNumber AS Invoice,
  vw_SixKeys.UnitCost AS Cost,
  vw_SixKeys.SchId,
  DrawDowns.DrawDownId
FROM
  DrawDowns
  INNER JOIN vw_SixKeys ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
WHERE
  (
    (
      (DrawDowns.DrawDownId)= [Forms]![DrawDownManager]![DrawDownId]
    )
  )
ORDER BY
  vw_SixKeys.Unitnum;

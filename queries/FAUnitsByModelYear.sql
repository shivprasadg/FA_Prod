SELECT
  vw_SixKeys.MYear AS ModelYear,
  vw_SixKeys.Make,
  vw_SixKeys.Type,
  Count(vw_SixKeys.UnitID) AS Units
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.MYear,
  vw_SixKeys.Make,
  vw_SixKeys.Type
ORDER BY
  vw_SixKeys.MYear DESC,
  vw_SixKeys.Make,
  vw_SixKeys.Type;

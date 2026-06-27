SELECT
  IIf(
    [Status] = "B",
    2016,
    DatePart("yyyy", [InServiceDate])
  ) AS YearInServiced,
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
  IIf(
    [Status] = "B",
    2016,
    DatePart("yyyy", [InServiceDate])
  )
ORDER BY
  IIf(
    [Status] = "B",
    2016,
    DatePart("yyyy", [InServiceDate])
  ) DESC;

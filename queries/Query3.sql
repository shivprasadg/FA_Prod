SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  Sum(Units.UnitIFamt) AS SumOfUnitIFamt,
  Min(Units.unitifdate) AS MinOfunitifdate,
  DateAdd(
    "m",
    6,
    Min([UnitIFDate])
  ) AS Expr1
FROM
  (
    Units
    INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
  )
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule;

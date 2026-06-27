SELECT
  UnitGroupPerdiemsSyndReport.GroupID,
  UnitGroupPerdiemsSyndReport.PDType,
  Avg(
    UnitGroupPerdiemsSyndReport.PerdiemRate
  ) AS PerdiemRate,
  Sum(
    IIf(
      [PerdiemDays] <= 0, 0, [PerdiemDays]
    )
  ) AS BilledPDDays,
  Sum(
    UnitGroupPerdiemsSyndReport.PerdiemTotal
  ) AS PerdiemTotal
FROM
  UnitGroupPerdiemsSyndReport
GROUP BY
  UnitGroupPerdiemsSyndReport.GroupID,
  UnitGroupPerdiemsSyndReport.PDType;

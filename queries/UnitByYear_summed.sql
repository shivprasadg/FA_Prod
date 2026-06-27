SELECT
  UnitsByYear.DYear,
  UnitsByYear.UnitType,
  Sum(
    (
      [A2014] + [A2015] + [A2016] + [A2017] + [A2018]
    )
  ) AS Active,
  Sum(
    (
      [New2014] + [New2015] + [New2016] + [New2017] + [New2018]
    )
  ) AS New,
  Sum(
    (
      [added2014] + [added2015] + [added2016] + [added2017] + [added2018]
    )
  ) AS Added
FROM
  UnitsByYear
GROUP BY
  UnitsByYear.DYear,
  UnitsByYear.UnitType
ORDER BY
  UnitsByYear.DYear,
  UnitsByYear.UnitType;

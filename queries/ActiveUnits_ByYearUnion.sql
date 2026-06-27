SELECT
  2013 AS Year,
  ClientShortName As Client,
  UnitType,
  Sum(A2012) AS YStart,
  Sum(D2013) As DUnits,
  Sum(O2013)*-1 As EOLUnits,
  Sum(A2013) AS ActiveUnits,
  Sum(N2013) As NewBus,
  Sum(XIT2013) As ExchIT,
  Sum(EYear2013) As ExtUnits
FROM
  ActiveUnits_ByYear
GROUP BY
  2013,
  ClientShortName,
  UnitType
Union All
SELECT
  2014 AS Year,
  ClientShortName,
  UnitType,
  Sum(A2013) AS YStart,
  Sum(D2014) As DUnits,
  Sum(O2014)*-1 As EOLUnits,
  Sum(A2014) AS ActiveUnits,
  Sum(N2014) As NewBus,
  Sum(XIT2014) As ExchIT,
  Sum(EYear2014) As ExtUnits
FROM
  ActiveUnits_ByYear
GROUP BY
  2014,
  ClientShortName,
  UnitType
Union All
SELECT
  2015 AS Year,
  ClientShortName,
  UnitType,
  Sum(A2014) AS YStart,
  Sum(D2015) As DUnits,
  Sum(O2015)*-1 As EOLUnits,
  Sum(A2015) AS ActiveUnits,
  Sum(N2015) As NewBus,
  Sum(XIT2015),
  Sum(EYear2015)
FROM
  ActiveUnits_ByYear
GROUP BY
  2015,
  ClientShortName,
  UnitType
Union All
SELECT
  2016 AS Year,
  ClientShortName,
  UnitType,
  Sum(A2015) AS YStart,
  Sum(D2016) As DUnits,
  Sum(O2016)*-1 As EOLUnits,
  Sum(A2016) AS ActiveUnits,
  Sum(N2016) As NewBus,
  Sum(XIT2016),
  Sum(EYear2016)
FROM
  ActiveUnits_ByYear
GROUP BY
  2016,
  ClientShortName,
  UnitType
Union All
SELECT
  2017 AS Year,
  ClientShortName,
  UnitType,
  Sum(A2016) AS YStart,
  Sum(D2017) As DUnits,
  Sum(O2017)*-1 As EOLUnits,
  Sum(A2017) AS ActiveUnits,
  Sum(N2017) As NewBus,
  Sum(XIT2017),
  Sum(EYear2017)
FROM
  ActiveUnits_ByYear
GROUP BY
  2017,
  ClientShortName,
  UnitType
UNION ALL
SELECT
  2018 AS Year,
  ClientShortName,
  UnitType,
  Sum(A2017) AS YStart,
  Sum(D2018) As DUnits,
  Sum(O2018)*-1 As EOLUnits,
  Sum(B2018)+ Sum(TR2018)+ Sum(A2018) AS ActiveUnits,
  Sum(N2018) As NewBus,
  Sum(XIT2018),
  Sum(EYear2018)
FROM
  ActiveUnits_ByYear
GROUP BY
  2018,
  ClientShortName,
  UnitType;

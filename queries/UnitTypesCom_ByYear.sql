SELECT
  UnitsTypesAccepted_ByYear.Accepted,
  UnitsTypesAccepted_ByYear.UnitType,
  Sum(
    UnitsTypesAccepted_ByYear.UnitAccepted
  ) AS SumOfUnitAccepted,
  UnitsTypesNewBusiness_ByYear.NewBusiness,
  UnitsTypesEOL_ByYear.UnitEOL
FROM
  (
    UnitsTypesAccepted_ByYear
    LEFT JOIN UnitsTypesNewBusiness_ByYear ON (
      UnitsTypesAccepted_ByYear.Accepted = UnitsTypesNewBusiness_ByYear.NewClient
    )
    AND (
      UnitsTypesAccepted_ByYear.UnitType = UnitsTypesNewBusiness_ByYear.UnitType
    )
  )
  LEFT JOIN UnitsTypesEOL_ByYear ON (
    UnitsTypesAccepted_ByYear.UnitType = UnitsTypesEOL_ByYear.UnitType
  )
  AND (
    UnitsTypesAccepted_ByYear.Accepted = UnitsTypesEOL_ByYear.EOL
  )
GROUP BY
  UnitsTypesAccepted_ByYear.Accepted,
  UnitsTypesAccepted_ByYear.UnitType,
  UnitsTypesNewBusiness_ByYear.NewBusiness,
  UnitsTypesEOL_ByYear.UnitEOL;

SELECT
  First(qry_FirstLastLoc.FirstCHdt1) AS FirstOfFirstCHdt1,
  Last(qry_FirstLastLoc.LastCHdt2) AS LastOfLastCHdt2,
  First(
    qry_FirstLastLoc.FirstOfOldUnitloc
  ) AS FirstOfFirstOfOldUnitloc,
  First(
    qry_FirstLastLoc.FirstOldStreet
  ) AS FirstOfFirstOldStreet,
  First(qry_FirstLastLoc.FirstOldCity) AS FirstOfFirstOldCity,
  First(
    qry_FirstLastLoc.FirstOldCounty
  ) AS FirstOfFirstOldCounty,
  First(qry_FirstLastLoc.FirstOldState) AS FirstOfFirstOldState,
  First(qry_FirstLastLoc.FirstOldZip) AS FirstOfFirstOldZip,
  First(
    qry_FirstLastLoc.FirstOldTaxStatus
  ) AS FirstOfFirstOldTaxStatus,
  First(
    qry_FirstLastLoc.FirstoldTaxNote
  ) AS FirstOfFirstoldTaxNote,
  Last(
    qry_FirstLastLoc.LastOfUnitsFormloc
  ) AS LastOfLastOfUnitsFormloc,
  Last(qry_FirstLastLoc.LastNewStreet) AS LastOfLastNewStreet,
  Last(
    qry_FirstLastLoc.LastNewTaxStatus
  ) AS LastOfLastNewTaxStatus,
  Last(
    qry_FirstLastLoc.LastNewTaxNote
  ) AS LastOfLastNewTaxNote,
  Last(qry_FirstLastLoc.NewLocation) AS LastOfNewLocation,
  Last(qry_FirstLastLoc.LastNewCounty) AS LastOfLastNewCounty,
  Last(qry_FirstLastLoc.LastNewState) AS LastOfLastNewState,
  Last(qry_FirstLastLoc.LastNewZip) AS LastOfLastNewZip,
  Last(qry_FirstLastLoc.MinRptChDate) AS LastOfMinRptChDate,
  Last(qry_FirstLastLoc.MaxRptChDate) AS LastOfMaxRptChDate,
  qry_FirstLastLoc.unitref
FROM
  qry_FirstLastLoc
GROUP BY
  qry_FirstLastLoc.unitref
HAVING
  (
    (
      (qry_FirstLastLoc.unitref)= 2359
    )
  );

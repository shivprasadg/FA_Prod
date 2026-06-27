SELECT
  vw_UnitVinDetails.unitid,
  vw_UnitVinDetails.VIN,
  vw_UnitVinDetails.modelyear,
  vw_UnitVinDetails.MAKE,
  vw_UnitVinDetails.MODEL
FROM
  vw_UnitVinDetails
WHERE
  (
    (
      (vw_UnitVinDetails.unitid) In (
        SELECT
          [UNITID]
        FROM
          [vw_UnitVinDetails] As Tmp
        GROUP BY
          [UNITID]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  vw_UnitVinDetails.unitid;

SELECT
  UMP.UnitID,
  UMP.LogActivityDate AS CurrentDate,
  UMP.Miles AS CurrentMiles,
  UMP.Fuel AS CurrentFuel,
  UMP.FuelMoving AS CurrentFuelMov,
  UMP.Hours AS CurrentHours,
  UMP.INDX
FROM
  vw_UnitMileageProcessed AS UMP
WHERE
  (
    (
      (UMP.[INDX]) In (
        SELECT
          MAX(INDX)
        FROM
          vw_UnitMileageProcessed AS UMP2
        GROUP BY
          UnitID
      )
    )
  );

SELECT
  vw_UnitMileageProcessed.UnitId AS UnitRef,
  Max(
    vw_UnitMileageProcessed.MilesLTD
  ) AS Mileage,
  IIf(
    [Fuel] > 0
    And [Mileage] > 0,
    Round([Mileage] / [Fuel], 2),
    0
  ) AS OverAllMPG,
  Max(
    vw_UnitMileageProcessed.FuelLTD
  ) AS Fuel,
  Max(
    vw_UnitMileageProcessed.LogActivityDate
  ) AS LastReport
FROM
  vw_UnitMileageProcessed
GROUP BY
  vw_UnitMileageProcessed.UnitId;

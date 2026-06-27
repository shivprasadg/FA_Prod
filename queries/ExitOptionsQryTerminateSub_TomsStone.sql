SELECT
  DISTINCTROW UnitExitOption.*
FROM
  UnitsFormFilter
  INNER JOIN UnitExitOption ON UnitsFormFilter.AssetIDx = UnitExitOption.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 2
    )
  );

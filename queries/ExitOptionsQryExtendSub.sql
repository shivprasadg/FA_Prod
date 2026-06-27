SELECT
  DISTINCTROW UnitExitOption.*,
  Units.UnitUnitNum
FROM
  UnitExitOption
  INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 3
    )
  );

SELECT
  UnitExitOption.UnitID,
  UnitExitOption.StartOrPurposedDate,
  UnitExitOption.OptionTypeID,
  Max(UnitExitOption.EndOrActualDate) AS EndOrActualDate,
  Sum(UnitExitOption.TermMonths) AS TotalMonths,
  Count(UnitExitOption.UnitId) AS TimesExtended,
  UnitExitOption.Fees
FROM
  UnitExitOption
GROUP BY
  UnitExitOption.UnitID,
  UnitExitOption.StartOrPurposedDate,
  UnitExitOption.OptionTypeID,
  UnitExitOption.Fees
HAVING
  (
    (
      (UnitExitOption.OptionTypeID)= 3
    )
  );

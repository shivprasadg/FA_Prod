SELECT
  UnitExitOption.UnitID,
  IIf([OptionTypeID] = 1, "Exch", Null) AS Exch,
  IIf([OptionTypeID] = 2, "Term", Null) AS Term,
  IIf([OptionTypeID] > 2, "Extd", Null) AS Extend
FROM
  UnitExitOption
GROUP BY
  UnitExitOption.UnitID,
  IIf([OptionTypeID] = 1, "Exch", Null),
  IIf([OptionTypeID] = 2, "Term", Null),
  IIf([OptionTypeID] > 2, "Extd", Null);

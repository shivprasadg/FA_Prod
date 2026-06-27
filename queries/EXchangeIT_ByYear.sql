SELECT
  UnitExitOption.ReplacedByGroupID,
  UnitExitOption.OptionTypeID,
  UnitExitOption.OptionTypeDesc,
  Count(UnitExitOption.UnitID) AS Units
FROM
  UnitExitOption
GROUP BY
  UnitExitOption.ReplacedByGroupID,
  UnitExitOption.OptionTypeID,
  UnitExitOption.OptionTypeDesc
HAVING
  (
    (
      (UnitExitOption.OptionTypeID)= 1
    )
  )
ORDER BY
  UnitExitOption.OptionTypeID;

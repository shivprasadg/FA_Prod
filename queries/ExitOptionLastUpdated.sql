SELECT
  UnitExitOption.UnitID,
  Max(UnitExitOption.DatePosted) AS MaxOfDatePosted,
  UnitExitOption.LeaseOptionID
FROM
  UnitExitOption
GROUP BY
  UnitExitOption.UnitID,
  UnitExitOption.LeaseOptionID;

SELECT
  LeaseOptionID,
  UO.UnitID,
  StartOrPurposedDate AS XStart,
  EndOrActualDate AS XEnd,
  DatePosted,
  OptionTypeID
FROM
  UnitExitOption AS UO
  INNER JOIN (
    SELECT
      Ex.UnitID,
      Max(Ex.EndOrActualDate) AS MaxEndDate
    FROM
      UnitExitOption AS Ex
    WHERE
      OptionTypeID > 2
    GROUP BY
      Ex.UnitID
  ) AS X ON (
    X.MaxEndDate = UO.EndOrActualDate
  )
  AND (UO.UnitID = X.UnitID);

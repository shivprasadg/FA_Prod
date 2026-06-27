SELECT
  PMs.UnitRef,
  Max(PMs.PMDate) AS MaxOfPMDate,
  Units.GroupREF
FROM
  PMs
  INNER JOIN Units ON PMs.UnitRef = Units.UnitID
GROUP BY
  PMs.UnitRef,
  Units.GroupREF;

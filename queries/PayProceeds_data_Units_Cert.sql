SELECT
  'Schedules:<br/>3-Units (3) 2023 Peterbilt 389 T/A Slpr Tract<br />' AS Schs,
  *
FROM
  vw_PayProceeds_UnitParts_Union
WHERE
  SchID in (1988)
  AND InterimAmt > 0;

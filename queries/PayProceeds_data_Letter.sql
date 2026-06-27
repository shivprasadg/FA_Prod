SELECT
  'Schedules:<br/>3-Units (3) 2023 Peterbilt 389 T/A Slpr Tract<br />' AS Schs,
  '3-Parts (1) Freight various<br />' AS VendorSchs,
  '1115502.79' AS TotalProceeds,
  *
FROM
  vw_PayProceeds_Summary
WHERE
  SchID in (1988);

SELECT
  *
FROM
  vw_CustodianTitleTrust2
WHERE
  uStatus <> 'C'
  and SchID = 2863
ORDER BY
  [Leasee Schedule No],
  [FA TitleLog Num];

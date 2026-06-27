SELECT
  ClientGroups.ClientGroupId,
  0 AS LocationID,
  "Not Required" AS LocationName,
  "Not Required" AS LAddrLine1,
  NULL AS City,
  NULL AS St,
  NULL AS Zip
FROM
  ClientGroups;

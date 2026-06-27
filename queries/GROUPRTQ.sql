SELECT
  ClientGroups.ClientGroupId,
  ClientGroups.ClientGroupName
FROM
  ClientGroups
GROUP BY
  ClientGroups.ClientGroupId,
  ClientGroups.ClientGroupName
ORDER BY
  ClientGroups.ClientGroupName;

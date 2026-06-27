SELECT
  ClientMasterLeasesTypeCount.clientID,
  ClientMasterLeasesTypeCount.Client,
  Switch(
    Max(
      [ClientMasterLeasesTypeCount].[FA]
    )= 1
    And Max(
      [ClientMasterLeasesTypeCount].[NotFA]
    )= 1,
    "Both",
    Max(
      [ClientMasterLeasesTypeCount].[FA]
    )= 1,
    "FA",
    Max(
      [ClientMasterLeasesTypeCount].[NotFA]
    )= 1,
    "Port"
  ) AS Type,
  Sum(
    ClientMasterLeasesTypeCount.Units
  ) AS Units,
  ClientMasterLeasesTypeCount.ClientGroupId
FROM
  ClientMasterLeasesTypeCount
GROUP BY
  ClientMasterLeasesTypeCount.clientID,
  ClientMasterLeasesTypeCount.Client,
  ClientMasterLeasesTypeCount.ClientGroupId
ORDER BY
  ClientMasterLeasesTypeCount.Client;

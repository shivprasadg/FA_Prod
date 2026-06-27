SELECT
  UnitHistory.AssetId,
  Max(UnitHistory.EventDate) AS EventDate,
  Last(UnitHistory.UnitHistoryId) AS UnitHistoryId,
  UnitHistory.EventGroup
FROM
  UnitHistory
GROUP BY
  UnitHistory.AssetId,
  UnitHistory.EventGroup;

SELECT
  LastTitleEntry.AssetID,
  LastTitleEntry.TitleLogId,
  LastTitleEntry.LogDate,
  LastTitleEntry.Action,
  LastTitleEntry.LastOfTitleNumber AS TitleNumber,
  LastTitleEntry.LastOfReason AS Description,
  LastTitleEntry.ExpectedReturn
FROM
  LastTitleEntry;

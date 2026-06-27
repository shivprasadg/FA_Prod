SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.BDEX,
  vw_SixKeys.BankShortName AS Assignee,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum AS UnitNum,
  vw_SixKeys.VIN,
  vw_SixKeys.Status,
  UnitTitleHistoryLastEntry.UnitTitleLogDate AS [Date Logged],
  UnitTitleHistoryLastEntry.EntryReason AS [Reason Out],
  UnitTitleHistoryLastEntry.TitleNumber AS [Title Number],
  UnitTitleHistoryLastEntry.TitledLocationState AS [Reg State],
  DateDiff(
    "d",
    [UnitTitleLogDate],
    Date()
  ) AS DaysOut,
  Units.LienRelease
FROM
  (
    UnitTitleHistoryLastEntry
    INNER JOIN vw_SixKeys ON UnitTitleHistoryLastEntry.UnitRef = vw_SixKeys.AssetID
  )
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
WHERE
  (
    (
      (
        UnitTitleHistoryLastEntry.ExpectedReturn
      )= True
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  UnitTitleHistoryLastEntry.UnitTitleLogDate;

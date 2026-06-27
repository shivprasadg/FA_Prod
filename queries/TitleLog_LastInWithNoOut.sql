SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Status,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  "IN" AS Direction,
  UnitTitleHistoryLastRecv.TitleNumber,
  UnitTitleHistoryLastRecv.EntryReason,
  UnitTitleHistoryLastRecv.UnitTitleLogDate
FROM
  (
    vw_SixKeys
    INNER JOIN UnitTitleHistoryLastRecv ON vw_SixKeys.AssetID = UnitTitleHistoryLastRecv.UnitRef
  )
  LEFT JOIN UnitTitleHistoryLastSent ON UnitTitleHistoryLastRecv.UnitRef = UnitTitleHistoryLastSent.UnitRef
WHERE
  (
    (
      (
        UnitTitleHistoryLastRecv.UnitTitleLogDate
      )> [UnitTitleHistoryLastSent].[UnitTitleLogDate]
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup;

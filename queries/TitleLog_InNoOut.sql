SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Status,
  LastTitleEntry_In.UnitTitleLogID,
  LastTitleEntry_In.UnitRef,
  LastTitleEntry_In.Action,
  LastTitleEntry_In.EntryReason,
  LastTitleEntry_In.TitleNumber,
  LastTitleEntry_In.UnitTitleLogDate
FROM
  vw_SixKeys
  INNER JOIN (
    LastTitleEntry_In
    LEFT JOIN LastTitleEntry_Out ON LastTitleEntry_In.UnitRef = LastTitleEntry_Out.UnitRef
  ) ON vw_SixKeys.AssetID = LastTitleEntry_In.UnitRef
WHERE
  (
    (
      (LastTitleEntry_Out.EntryReason) Is Null
    )
    AND (
      (
        LastTitleEntry_Out.UnitTitleLogDate
      ) Is Null
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup;

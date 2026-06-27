SELECT
  vw_SixKeys.AssetID
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.SchID)= [Forms]![ScheduleControl]![SchID]
    )
  )
ORDER BY
  vw_SixKeys.AssetID;

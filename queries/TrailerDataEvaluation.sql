SELECT
  TrailerMilesHours.UnitID,
  Max(
    MainesTrailersDataSource.ReadDateMiles
  ) AS ReadDate,
  Max(
    TrailerMilesHours.LogActivityDate
  ) AS LastDate,
  Max(
    MainesTrailersDataSource.LogActivityDate
  ) AS NewDate,
  Max(TrailerMilesHours.MilesLTD) AS PrevMilesLTD,
  Max(MainesTrailersDataSource.Miles) AS CurrMilesLTD,
  [CurrMilesLTD] - [PrevMilesLTD] AS MilesDriven,
  DateDiff("m", [LastDate], [NewDate]) AS MonthDiff,
  ([MonthDiff] * 14000) AS MilesLimit,
  Switch(
    [MilesDriven] = 0, 0, [MilesDriven] > [MilesLimit],
    2, [MilesDriven] < 0, 3, [MilesDriven] > 0,
    1
  ) AS MilesDataStatus,
  Switch(
    [MilesDataStatus] = 2,
    "Excess Miles: " & ([MilesDriven] - [MilesLimit]),
    [MilesDataStatus] = 3,
    "Neg Miles",
    [MilesDataStatus] = 0,
    "No Miles",
    [MilesDataStatus] = 1,
    "Reported Miles"
  ) AS MilesDataStatusDesc,
  IIf(
    [MilesDriven] <= 0, 0, [MilesDriven]
  ) AS MilesResult,
  Max(TrailerMilesHours.HoursLTD) AS PrevHoursLTD,
  Max(MainesTrailersDataSource.Hours) AS CurrHoursLTD,
  [CurrHoursLTD] - [PrevHoursLTD] AS HoursUsed,
  [MonthDiff] *(24 * 30) AS HoursLimit,
  Switch(
    [HoursUsed] > [HoursLimit], 2, [HoursUsed] < 0,
    3, [HoursUsed] = 0, 0, [HoursUsed] > 0,
    1
  ) AS HoursDataStatus,
  Switch(
    [HoursDataStatus] = 2, "Excess Hours: " & [HoursUsed] - [HoursLimit],
    [HoursDataStatus] = 3, "Neg Hours",
    [HoursDataStatus] = 0, "No Hours",
    [HoursDataStatus] = 1, "Reported Hours"
  ) AS HoursDataStatusDesc,
  [MilesDataStatusDesc] & ", " & [HoursDataStatusDesc] AS ObcDataDesc,
  IIf(
    [ObcDataDesc] = "Reported Miles, Reported Hours",
    "Reporting", [ObcDataDesc]
  ) AS DataDesc,
  IIf(
    [MilesDataStatus] + [HoursDataStatus] Between 1
    And 2,
    1, [MilesDataStatus] + [HoursDataStatus]
  ) AS OBCDataStatus
FROM
  MainesTrailersDataSource
  LEFT JOIN TrailerMilesHours ON MainesTrailersDataSource.UnitID = TrailerMilesHours.UnitID
WHERE
  (
    (
      (
        MainesTrailersDataSource.LogActivityDate
      )> [TrailerMilesHours].[LogActivityDate]
    )
  )
GROUP BY
  TrailerMilesHours.UnitID;

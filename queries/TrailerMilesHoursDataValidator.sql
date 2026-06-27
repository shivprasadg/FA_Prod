SELECT
  TrailerMilesHours.UnitID,
  TrailerMilesHours.UnitNumber,
  TrailerMilesHours.LogActivityDate,
  Switch(
    [Miles] = 0, 0, [Miles] >= 1
    And [Miles] < 35000,
    1, [Miles] < 0, 16, [Miles] >= 35000,
    14
  ) AS MilesCode,
  Switch(
    [MilesCode] = 14, "Excess Miles", [MilesCode] = 16,
    "Neg Miles", [MilesCode] = 0, "No Miles",
    [MilesCode] = 1, "Good Miles"
  ) AS MilesCodeDesc,
  Switch(
    [Hours] > 1000, 22, [Hours] < 0, 33, [Hours] = 0,
    0, [Hours] >= 1
    And [Hours] < 1000, 3
  ) AS HoursCode,
  Switch(
    [HoursCode] = 22, "Excess Hours", [HoursCode] = 33,
    "Neg Hours", [HoursCode] = 0, "No Hours",
    [HoursCode] = 3, "Good Hours"
  ) AS HoursCodeDesc,
  IIf(
    [MilesCode] + [HoursCode] < 5, 1, [MilesCode] + [HoursCode]
  ) AS DataStatus,
  Switch(
    [MilesCode] + [HoursCode] = 0, "Reported - No Update",
    [MilesCode] + [HoursCode] = 1, "Reported - No Hours",
    [MilesCode] + [HoursCode] = 3, "Reported - No Miles",
    [MilesCode] + [HoursCode] = 4, "Reporting",
    [MilesCode] + [HoursCode] > 4, [MilesCodeDesc] & "-" & [HoursCodeDesc]
  ) AS DataStatusDesc
FROM
  TrailerMilesHours;

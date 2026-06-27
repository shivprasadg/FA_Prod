SELECT
  Units.UnitID,
  Units.UnitYr,
  UnitExitOption.OptionTypeID,
  CCur(
    Nz([unitrent], [Rent])
  ) AS ActvRent,
  DateAdd("m", [LeaseTermPrimary], [BLCD])-1 AS PriEOL,
  UnitExitOption.EndOrActualDate AS ExtendedEOL,
  Nz([EndOrActualDate], [PriEOL]) AS ProjEOL,
  Units.unitoffleasedt AS ActualEOL,
  IIf(
    IsNull([unitoffleasedt]),
    IIf(
      IsNull([EndOrActualDate]),
      "Primary",
      "Extended"
    ),
    "Actual"
  ) AS EOLType,
  IIf(
    IsNull([ActualEOL]),
    DateDiff("m", [PriEOL], [ProjEOL]),
    DateDiff("m", [PriEOL], [ActualEOL])
  ) AS MonthsDiff,
  Switch(
    [MonthsDiff] < 0,
    "Early:" & Abs([MonthsDiff]),
    [MonthsDiff] = 0,
    "None",
    [MonthsDiff] > 0,
    "GAP: " & Abs([MonthsDiff])
  ) AS GAPType
FROM
  (
    Units
    LEFT JOIN UnitExitOption ON Units.UnitID = UnitExitOption.UnitID
  )
  INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 3
      Or (UnitExitOption.OptionTypeID) Is Null
    )
  )
ORDER BY
  Units.UnitID;

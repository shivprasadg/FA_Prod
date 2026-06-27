SELECT
  vw_SixKeys.ClientShortName AS Client,
  "Sch " & [Schedule] & " - Grp " & [UnitGroup] AS Fleet_Advantage_Agreement,
  Nz([SchBnkClNO], "")& IIf(
    Nz([SgrpBnkNo], "")<> "",
    "-" & [SgrpBnkNo],
    ""
  ) AS [Assignee Contract No],
  vw_SixKeys.BankShortName AS [Assignee Bank],
  vw_SixKeys.UnitNum AS [Client Unit Num],
  vw_SixKeys.VIN,
  vw_SixKeys.UnitDescDocVer AS UnitDesc,
  [LCity] & " " & [LState] AS Location,
  CCur([UnitCost]) AS [Unit Cost],
  CCur([PartsCost]) AS [Child Cost],
  [UnitCost] + [PartsCost] AS [Total Cost],
  vw_SixKeys.BLCD AS Commencement_Date,
  vw_SixKeys.AcceptedDate,
  Lookup_LeaseStartsOn.StartsOn AS PerDiemStart,
  Units.UnitPDStartDate AS PerDiemStartDate,
  vw_SixKeys.InServiceDate,
  vw_SixKeys.BLCD,
  IIf(
    IsDate([UnitPDStartDate])
    And IsDate([vw_SixKeys].[BLCD]),
    DateDiff(
      "d", [UnitPDStartDate], [vw_SixKeys].[BLCD]
    ),
    0
  ) AS PDDays,
  Units.UNITPDRENT AS PDRent,
  Nz([PDDays], 0)* Nz([UnitPDRent], 0) AS EstPerdiemTotal,
  vw_SixKeys.ProjEOL AS Final_Payment_Date,
  Units.unitsrent AS SchRent,
  Units.unitrent AS SyndRent,
  FormatPercent(
    IIf(
      [UnitRent] > 0,
      Round([UnitRent] / [Total Cost], 7),
      0
    ),
    4
  ) AS [Lease_Rate Factor],
  [SchGrp].[LeaseTermPrimary] - Nz([SchGrp].[LeaseTermNotice], 0) AS [Primary Lease Term],
  Nz([SchGrp].[LeaseTermNotice], 0) AS NoticeReq,
  IIf(
    IsDate([OffleaseActual]),
    0,
    DateDiff(
      "m",
      Date(),
      [ProjEOL]
    )
  ) AS MthsLeft,
  IIf(
    IsDate([OffleaseActual]),
    0,
    DateDiff(
      "m",
      Date(),
      [ProjEOL]
    )* [CurrRent]
  ) AS [Remain Rent],
  Units.UnitActualDelDate AS Delivered,
  vw_SixKeys.SchID,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.Status
FROM
  (
    (
      (
        vw_SixKeys
        INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
      )
      INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
    )
    INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
  )
  LEFT JOIN Lookup_LeaseStartsOn ON SchGrp.PerDiemStartsOn = Lookup_LeaseStartsOn.ID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  "Sch " & [Schedule] & " - Grp " & [UnitGroup],
  [LCity] & " " & [LState];

SELECT
  E.Client,
  E.Schedule,
  E.UnitGroup,
  E.Assignee,
  E.LeaseType,
  E.BLCD,
  E.AssetID,
  E.UnitNum,
  E.VIN,
  U.UnitYr AS [Year],
  E.UnitMake AS Make,
  E.UnitModel AS Model,
  E.UnitSubType AS SubType,
  E.UnitType AS Type,
  E.CurrRent,
  E.CurrRentType,
  E.uStatus,
  E.EOLPrimary AS [Primary Exp],
  E.DaysToEOL,
  E.Selected AS [Exit Options],
  E.ExtendedStart AS ExtdStart,
  E.OffLeaseProj AS ExtdEnd,
  E.SufSent,
  E.SufBack,
  DateDiff(
    "m",
    [SufSent],
    Nz(
      [SUFBack],
      Date()
    )
  ) AS SentAge,
  DateDiff(
    "m",
    [SufBack],
    Date()
  ) AS BackAge,
  vw_UnitInspectionLast.iDate,
  E.iPassedDate,
  E.ActualOffLease,
  B.LeaseRV,
  B.BookNow,
  B.BV90Days,
  E.ODORead,
  E.ODOSource,
  E.ODOReadDate,
  E.Garage,
  E.ExchDate,
  E.NewClient,
  E.NewSchedule,
  E.NewGroup,
  E.NewAssetID,
  E.nInServ,
  E.ClientID,
  E.SchID,
  E.GroupID,
  LastExtendedEvent.LeaseOptionID
FROM
  (
    (
      (
        ExitOptionsGroup_Local AS E
        INNER JOIN Units AS U ON E.AssetID = U.UnitID
      )
      LEFT JOIN vw_BookValue AS B ON E.AssetID = B.AssetID
    )
    LEFT JOIN vw_UnitInspectionLast ON E.AssetID = vw_UnitInspectionLast.AssetID
  )
  LEFT JOIN LastExtendedEvent ON U.UnitID = LastExtendedEvent.UnitID
WHERE
  (
    (
      Not (E.SufSent) Is Null
    )
    AND (
      (E.ActualOffLease) Is Null
    )
  );

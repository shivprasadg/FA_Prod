SELECT
  DISTINCT Clients.ClientShNm AS Client,
  Deals.dealShortDesc AS DealDesc,
  Deals.dealOutcome AS DealResult,
  Min(Deals.dealEnteredDate) AS DealEntered,
  Deals.dealOutcomeDate AS ResultDate,
  DateDiff(
    'd',
    Min([dealEnteredDate]),
    [dealOutcomeDate]
  ) AS DaysToAward,
  Min(
    Schedule_First_BLCDate.Schedule
  ) AS Schedule,
  IIf(
    IsNull(
      Min([DateSchAdded])
    ),
    "PENDING",
    Min([DateSchAdded])
  ) AS ScheduleAdded,
  IIf(
    IsNull(
      Min([DateSchAdded])
    ),
    DateDiff(
      'd',
      [dealOutcomeDate],
      Date()
    ),
    DateDiff(
      'd',
      [dealOutcomeDate],
      Min([DateSchAdded])
    )
  ) AS LPAFDays,
  Min(
    Schedule_First_BLCDate.Assignee
  ) AS Assignee,
  Nz(
    Min([FirstBLCD]),
    "PENDING"
  ) AS BLCD,
  IIf(
    IsNull(
      Min([FirstBLCD])
    ),
    DateDiff(
      'd',
      Min([DateSchAdded]),
      Date()
    ),
    DateDiff(
      'd',
      Min([DateSchAdded]),
      Min([FirstBLCD])
    )
  ) AS OPSDays,
  Nz([ATFMaxRevNo].[RevDt], "n/a") AS ATFRev
FROM
  (
    (
      Deals
      INNER JOIN Clients ON Deals.clientRef = Clients.clientID
    )
    LEFT JOIN Schedule_First_BLCDate ON (
      Deals.clientRef = Schedule_First_BLCDate.ClientID
    )
    AND (
      Deals.FAID = Schedule_First_BLCDate.FAID
    )
  )
  LEFT JOIN ATFMaxRevNo ON Schedule_First_BLCDate.ScheduleID = ATFMaxRevNo.SchID
GROUP BY
  Clients.ClientShNm,
  Deals.dealShortDesc,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  Nz([ATFMaxRevNo].[RevDt], "n/a")
HAVING
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
    AND (
      (
        Min(Deals.dealEnteredDate)
      )> #12/31/2012#
    )
  )
ORDER BY
  Clients.ClientShNm,
  Min(Deals.dealEnteredDate) DESC;

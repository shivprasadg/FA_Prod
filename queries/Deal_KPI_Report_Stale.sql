SELECT
  DISTINCT Clients.ClientShNm AS Client,
  Deals.dealShortDesc AS DealDesc,
  Deals.dealOutcome AS DealResult,
  Deals.dealEnteredDate AS DealEntered,
  Deals.dealOutcomeDate AS ResultDate,
  DateDiff(
    'd', [dealEnteredDate], [dealOutcomeDate]
  ) AS DaysToAward,
  First(
    Schedule_First_BLCDate.Schedule
  ) AS Schedule,
  Schedule_First_BLCDate.DateSchAdded AS ScheduleAdded,
  DateDiff(
    'd',
    [dealOutcomeDate],
    Date()
  ) AS StaleDays,
  First(
    Schedule_First_BLCDate.Assignee
  ) AS Assignee,
  First(
    Schedule_First_BLCDate.FirstBLCD
  ) AS BLCD,
  DateDiff(
    'd',
    [DateSchAdded],
    First([FirstBLCD])
  ) AS OPSDays
FROM
  (
    Deals
    LEFT JOIN Clients ON Deals.clientRef = Clients.clientID
  )
  LEFT JOIN Schedule_First_BLCDate ON Clients.clientID = Schedule_First_BLCDate.ClientID
WHERE
  (
    (
      (
        Schedule_First_BLCDate.ClientID
      ) Is Null
    )
  )
GROUP BY
  Clients.ClientShNm,
  Deals.dealShortDesc,
  Deals.dealOutcome,
  Deals.dealEnteredDate,
  Deals.dealOutcomeDate,
  Schedule_First_BLCDate.DateSchAdded,
  Deals.dealEnteredDate
HAVING
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
    AND (
      (Deals.dealEnteredDate)> #12/31/2012#
    )
  )
ORDER BY
  Clients.ClientShNm,
  Deals.dealEnteredDate DESC;

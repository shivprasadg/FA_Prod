SELECT
  DISTINCT Clients.ClientShNm AS Client,
  Deals.dealShortDesc AS DealDesc,
  Deals.dealOutcome AS DealResult,
  Max(
    DateValue([dealEnteredDate])
  ) AS DealEntered,
  Deals.dealOutcomeDate AS ResultDate,
  DateDiff(
    'd',
    Min([dealEnteredDate]),
    [dealOutcomeDate]
  ) AS DaysToAward,
  Min(
    Schedule_Group_Entered_KPI.Schedule
  ) AS Schedule,
  IIf(
    IsNull(
      Min([DateSchAdded])
    ),
    "PENDING",
    Min([DateSchAdded])
  ) AS ScheduleAdded,
  Min(
    Schedule_Group_Entered_KPI.Group
  ) AS [Group],
  IIf(
    IsNull(
      Min([DateGroupAdded])
    ),
    "PENDING",
    Min([DateGroupAdded])
  ) AS GroupAdded,
  IIf(
    IsNull(
      Min([DateGroupAdded])
    ),
    DateDiff(
      'd',
      [dealOutcomeDate],
      Date()
    ),
    DateDiff(
      'd',
      [dealOutcomeDate],
      Min([DateGroupAdded])
    )
  ) AS LPAFDays,
  Nz(
    Min(
      [Schedule_Group_Entered_KPI].[BLCD]
    ),
    "PENDING"
  ) AS BLCD,
  IIf(
    IsNull(
      Min(
        [Schedule_Group_Entered_KPI].[ApprovalRequest]
      )
    ),
    DateDiff(
      'd',
      Min([DateSchAdded]),
      Date()
    ),
    DateDiff(
      'd',
      Min([DateSchAdded]),
      Min(
        [Schedule_Group_Entered_KPI].[ApprovalRequest]
      )
    )
  ) AS OPSDays,
  Deals.FAID,
  Deals.dealUnitQty,
  Deals.dealGrandTotal,
  Schedule_Group_Entered_KPI.ApprovalRequest,
  Schedule_Group_Entered_KPI.ApprovalReply,
  Deals.ExcludeFromKPI
FROM
  (
    Deals
    INNER JOIN Clients ON Deals.clientRef = Clients.clientID
  )
  LEFT JOIN Schedule_Group_Entered_KPI ON (
    Deals.clientRef = Schedule_Group_Entered_KPI.ClientID
  )
  AND (
    Deals.FAID = Schedule_Group_Entered_KPI.FAID
  )
GROUP BY
  Clients.ClientShNm,
  Deals.dealShortDesc,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  Deals.FAID,
  Deals.dealUnitQty,
  Deals.dealGrandTotal,
  Schedule_Group_Entered_KPI.ApprovalRequest,
  Schedule_Group_Entered_KPI.ApprovalReply,
  Deals.ExcludeFromKPI
HAVING
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
    AND (
      (
        Max(
          DateValue([dealEnteredDate])
        )
      )> #12/31/2012#
    )
    AND (
      (Deals.ExcludeFromKPI) Is Null
      Or (Deals.ExcludeFromKPI)= 0
    )
  )
ORDER BY
  Clients.ClientShNm,
  Max(
    DateValue([dealEnteredDate])
  ) DESC;

SELECT
  SyndGroupStripsUTA.ScheduleID,
  SyndGroupStripsUTA.GroupID,
  Max(SyndGroupStripsUTA.StripsNum) AS GroupStrips,
  Sum(
    CCur(
      Nz([SalesRent], 0)* Nz([StripsNum], 0)
    )
  ) AS GroupStripAmount,
  Max(SyndGroupStripsUTA.UTA) AS UTAPer,
  Sum(SyndGroupStripsUTA.UTA) AS UTAGroupTotal
FROM
  SyndGroupStripsUTA
GROUP BY
  SyndGroupStripsUTA.ScheduleID,
  SyndGroupStripsUTA.GroupID;

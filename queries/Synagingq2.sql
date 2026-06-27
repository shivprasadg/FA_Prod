SELECT
  Synagingq1.ClientShNm AS Client,
  Synagingq1.MLNo,
  Synagingq1.SchNo AS [Sch No],
  Synagingq1.SGrpNO AS [Grp No],
  Synagingq1.BLCD AS BLCD,
  Synagingq1.AssignmentDate AS [Bank Assign Dt],
  Synagingq1.UnitsInGroup AS [No Of Units],
  IIf(
    IsNull([INserVIN]),
    0,
    [inserVIN]
  ) AS [Units In-Serv],
  Synagingq1.MaxOfunitacceptdate AS [Max Accepteded Dt],
  Date()- [maxofunitacceptdate] + 1 AS [Age Accepted],
  DateAdd("d", 90, [maxofunitacceptdate]) AS [AccpetDt+90],
  IIf(
    IsNull([inservin])
    Or [UnitsInGroup] <> [inservin],
    "TBD",
    [maxofunitinserv]
  ) AS [Max Inservice dt],
  Date()- [maxofunitinserv] + 1 AS [Age In-Service],
  DateAdd("d", 90, [maxofunitinserv]) AS [In-Inserv+90],
  Date() AS [As of],
  Clients.ClientGroupID
FROM
  Clients
  INNER JOIN (
    Synagingq1
    LEFT JOIN Synagingq1a ON Synagingq1.SchID = Synagingq1a.SchID
  ) ON Clients.ClientShNm = Synagingq1.ClientShNm
WHERE
  (
    (
      (Synagingq1.UnitsInGroup)= [CountOfunitvin]
    )
    AND (
      (Clients.ClientGroupID) Not Like "* demo"
    )
  )
ORDER BY
  Date()- [maxofunitinserv] + 1 DESC,
  Synagingq1.ClientShNm;

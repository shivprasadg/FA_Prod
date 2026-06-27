SELECT
  Salesperson.SPInitials AS SM,
  Clients.clientCompanyName AS [Co Name],
  Clients.clientCity AS City,
  Clients.clientState AS ST,
  Clients.clientZip AS Zip,
  Clients.fleetSize,
  Clients.fleetDaycabs AS [No of DC],
  Clients.fleetDaycabsMPY AS [DC MPY],
  Clients.fleetSleepers AS [No of Slp],
  Clients.fleetSleepersMPY AS SlpMPY,
  Clients.FleetReplacementCycle AS [Repl Cycle],
  Clients.creditType,
  Clients.creditMoodyRating AS Moodys,
  Clients.creditRank AS [FA Rank],
  [SGRD] + [DCGRD] + [DCMIGRD] + [SPLMIGRD] + [MAINTGRD] + [LCGRD] + [OBCGRD] + [MODGRD] + IIf(
    [creditRank] = "tba", [MOODYGRD], [FACRGRD]
  ) AS nscore,
  Clients.clienttype
FROM
  (
    Salesperson
    INNER JOIN Clients ON Salesperson.SalespersonID = Clients.SalespersonREF
  )
  INNER JOIN NSCOREQ ON Clients.clientID = NSCOREQ.clientID
WHERE
  (
    (
      (Clients.clienttype)= "target"
    )
  )
ORDER BY
  Clients.clientCompanyName;

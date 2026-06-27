SELECT
  IIf(
    [CLIENTTYPE] = "TARGET",
    1,
    IIf([CLIENTTYPE] = "2ND TARGET", 2, 3)
  ) AS SRTORDER,
  Clients.clienttype,
  Clients.loinitial AS SM,
  Clients.clientCompanyName AS [Co Name],
  Clients.clientCity AS City,
  Clients.clientState AS ST,
  Clients.clientZip AS Zip,
  Clients.fleetSize,
  Clients.fleetDaycabs AS [No of DC],
  NSCOREQ.DCMPY,
  NSCOREQ.DCMIO,
  Clients.fleetSleepers AS [No of Slp],
  NSCOREQ.SlpMPY,
  NSCOREQ.SLPMIO,
  Clients.FleetReplacementCycle AS [Repl Cycle],
  Clients.creditType,
  Clients.creditMoodyRating AS Moodys,
  Clients.creditRank AS [FA Rank],
  [SGRD] + [DCMIGRD] + [SPLMIGRD] + [MAINTGRD] + [LCGRD] + [OBCGRD] + [MODGRD] + IIf(
    [creditRank] = "tba", [MOODYGRD], [FACRGRD]
  ) AS nscore,
  Clients.creditFinance,
  Clients.creditMaint,
  Clients.clientobc,
  Clients.fleetDaycabsMPY,
  Clients.fleetSleepersMPY,
  [clients].[fleetdaycabs] + [clients].[fleetsleepers] AS FS,
  GROUPRTQ.maxofcreditfascore AS GRPMAX,
  Clients.ClientGroupID,
  Clients.JFApproval,
  IIf(
    [clients].[fleetdaycabs] + [clients].[fleetsleepers] = 0
    Or (
      [clients].[fleetdaycabs] > 0
      And [clients].[fleetdaycabsmpy] = 0
    )
    Or (
      [clients].[fleetsleepers] > 0
      And [clients].[fleetsleepersmpy] = 0
    )
    Or [clientobc] = "tba"
    Or [creditfinance] = "tbd"
    Or [creditmaint] = "tbd"
    Or IsNull([fleetreplacementcycle])
    Or (
      [creditrank] = "tb"
      And [creditmoodyrating] = "tba"
      And IsNull([jfapproval])
    ),
    True,
    0
  ) AS Mdata,
  TQ2.ex,
  TQ2.DT,
  Clients.creditRank,
  Clients.creditFAProfile,
  Clients.creditFAScore,
  Clients.creditFAScore
FROM
  (
    (
      (
        Clients
        LEFT JOIN NSCOREQ ON Clients.clientID = NSCOREQ.clientID
      )
      LEFT JOIN GROUPRTQ ON Clients.ClientGroupID = GROUPRTQ.ClientGroupID
    )
    LEFT JOIN TQ2 ON Clients.clientID = TQ2.clientID
  )
  INNER JOIN Salesperson ON Clients.LOinitial = Salesperson.SPInitials
GROUP BY
  IIf(
    [CLIENTTYPE] = "TARGET",
    1,
    IIf([CLIENTTYPE] = "2ND TARGET", 2, 3)
  ),
  Clients.clienttype,
  Clients.loinitial,
  Clients.clientCompanyName,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  Clients.fleetSize,
  Clients.fleetDaycabs,
  NSCOREQ.DCMPY,
  NSCOREQ.DCMIO,
  Clients.fleetSleepers,
  NSCOREQ.SlpMPY,
  NSCOREQ.SLPMIO,
  Clients.FleetReplacementCycle,
  Clients.creditType,
  Clients.creditMoodyRating,
  Clients.creditRank,
  [SGRD] + [DCMIGRD] + [SPLMIGRD] + [MAINTGRD] + [LCGRD] + [OBCGRD] + [MODGRD] + IIf(
    [creditRank] = "tba", [MOODYGRD], [FACRGRD]
  ),
  Clients.creditFinance,
  Clients.creditMaint,
  Clients.clientobc,
  Clients.fleetDaycabsMPY,
  Clients.fleetSleepersMPY,
  [clients].[fleetdaycabs] + [clients].[fleetsleepers],
  GROUPRTQ.maxofcreditfascore,
  Clients.ClientGroupID,
  Clients.JFApproval,
  IIf(
    [clients].[fleetdaycabs] + [clients].[fleetsleepers] = 0
    Or (
      [clients].[fleetdaycabs] > 0
      And [clients].[fleetdaycabsmpy] = 0
    )
    Or (
      [clients].[fleetsleepers] > 0
      And [clients].[fleetsleepersmpy] = 0
    )
    Or [clientobc] = "tba"
    Or [creditfinance] = "tbd"
    Or [creditmaint] = "tbd"
    Or IsNull([fleetreplacementcycle])
    Or (
      [creditrank] = "tb"
      And [creditmoodyrating] = "tba"
      And IsNull([jfapproval])
    ),
    True,
    0
  ),
  TQ2.ex,
  TQ2.DT,
  Clients.creditRank,
  Clients.creditFAProfile,
  Clients.creditFAScore,
  Clients.creditFAScore,
  Clients.ClientGroupID
HAVING
  (
    (
      (Clients.clienttype) Like "*target"
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
  )
ORDER BY
  Clients.loinitial,
  Clients.clientCompanyName;

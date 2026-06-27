SELECT
  Client_Lease_Detail.Client,
  Client_Lease_Detail.Fleet_Advantage_Agreement,
  Client_Lease_Detail.[Assignee Contract No],
  Client_Lease_Detail.[Assignee Bank],
  Client_Lease_Detail.UnitDesc,
  Client_Lease_Detail.Location,
  Count(Client_Lease_Detail.VIN) AS Units,
  Client_Lease_Detail.Commencement_Date,
  Client_Lease_Detail.Final_Payment_Date,
  IIf(
    Sum([SyndRent])> 0,
    Round(
      Sum([SyndRent])/ Sum([Total Cost]),
      5
    ),
    0
  ) AS [Lease Rate Factor],
  Sum(Client_Lease_Detail.SyndRent) AS [Estimated Rent],
  Client_Lease_Detail.ClientGroupId,
  Client_Lease_Detail.SchID
FROM
  Client_Lease_Detail
GROUP BY
  Client_Lease_Detail.Client,
  Client_Lease_Detail.Fleet_Advantage_Agreement,
  Client_Lease_Detail.[Assignee Contract No],
  Client_Lease_Detail.[Assignee Bank],
  Client_Lease_Detail.UnitDesc,
  Client_Lease_Detail.Location,
  Client_Lease_Detail.Commencement_Date,
  Client_Lease_Detail.Final_Payment_Date,
  Client_Lease_Detail.ClientGroupId,
  Client_Lease_Detail.SchID
ORDER BY
  Client_Lease_Detail.SchID;

SELECT
  DLookUp(
    "ATFID", "Schedule", "Schid=" & [vw_SixKeys].[Schid]
  ) AS ATFID,
  vw_SixKeys.FAID,
  vw_SixKeys.MLID,
  vw_SixKeys.ClientId AS LesseeID,
  1 AS CreditRank,
  DLookUp(
    "ATFSchSpecID", "Schedule", "Schid=" & [vw_SixKeys].[schid]
  ) AS ATFSchSpecID,
  vw_SixKeys.BankShortName AS Assignee_Bank,
  Min(
    "FAID: " & [vw_SixKeys].[FAID] & "-" & [dealEnteredDate]
  ) AS ATFDesc,
  vw_SixKeys.CompanyName AS Lessee,
  Count(vw_SixKeys.UnitID) AS ActualUnits,
  Min(Deals.dealEnteredDate) AS ATFDate,
  Min(Deals.dealLPAFcomp) AS LPAFDate,
  Deals.DealLOEx AS BDEX,
  vw_SixKeys.MLNo,
  DLookUp(
    "clientTaxID", "Clients", "ClientID=" & [vw_SixKeys].[ClientID]
  ) AS ClientTaxID,
  0 AS CoLesseeID,
  vw_SixKeys.BankId,
  ECount(
    "SchId", "vw_SixKeys", "BankId= " & [vw_SixKeys].[BankId] & " AND FAID='" & [vw_Sixkeys].[FAID] & "'",
    True
  ) AS NumOfSchedules,
  ECount(
    "SGrpId", "vw_SixKeys", "BankId= " & [vw_SixKeys].[BankId] & " AND FAID='" & [vw_Sixkeys].[FAID] & "'",
    True
  ) AS NumOfGroups,
  vw_SixKeys.SchID,
  Deals.dealOutcome
FROM
  vw_SixKeys
  LEFT JOIN Deals ON vw_SixKeys.FAID = Deals.FAID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  DLookUp(
    "ATFID", "Schedule", "Schid=" & [vw_SixKeys].[Schid]
  ),
  vw_SixKeys.FAID,
  vw_SixKeys.MLID,
  vw_SixKeys.ClientId,
  1,
  DLookUp(
    "ATFSchSpecID", "Schedule", "Schid=" & [vw_SixKeys].[schid]
  ),
  vw_SixKeys.BankShortName,
  vw_SixKeys.CompanyName,
  Deals.DealLOEx,
  vw_SixKeys.MLNo,
  DLookUp(
    "clientTaxID", "Clients", "ClientID=" & [vw_SixKeys].[ClientID]
  ),
  0,
  vw_SixKeys.BankId,
  ECount(
    "SchId", "vw_SixKeys", "BankId= " & [vw_SixKeys].[BankId] & " AND FAID='" & [vw_Sixkeys].[FAID] & "'",
    True
  ),
  ECount(
    "SGrpId", "vw_SixKeys", "BankId= " & [vw_SixKeys].[BankId] & " AND FAID='" & [vw_Sixkeys].[FAID] & "'",
    True
  ),
  vw_SixKeys.SchID,
  Deals.dealOutcome;

SELECT
  TBL_Employees.ID,
  TBL_Employees.[Employee Name] AS SalesPerson,
  TBL_Employees.EmpInt AS SP,
  TBL_Employees.[User ID],
  TBL_Employees.[Access Level],
  CInt(
    Nz(
      [Events_ByType_Last30].[callsIn],
      0
    )
  ) AS CallsIn,
  CInt(
    Nz(
      [Events_ByType_Last30].[CallsOut],
      0
    )
  ) AS CallsOut,
  CInt(
    Nz(
      [Events_ByType_Last30].[EmailsIn],
      0
    )
  ) AS EmailsIn,
  CInt(
    Nz(
      [Events_ByType_Last30].[EmailsOut],
      0
    )
  ) AS EmailsOut,
  CInt(
    Nz(
      [Events_ByType_Last30].[VisitsScheduled],
      0
    )
  ) AS VisitsScheduled,
  CInt(
    Nz(
      [BuyersProposalstatus_Last30].[Proposals],
      0
    )
  ) AS Proposals,
  CInt(
    Nz(
      [BuyersProposalstatus_last30].[ProposalActive],
      0
    )
  ) AS ProposalsActive,
  CInt(
    Nz(
      [BuyersProposalstatus_Last30].[Proposalsold],
      0
    )
  ) AS ProposalsSold,
  CInt(
    Nz([UnitsSoldLast30], 0)
  ) AS Sold30,
  Format(
    Nz([AmtSold30], 0),
    "$#,###"
  ) AS Sales30,
  CInt(
    Nz([UnitsSoldLast60], 0)
  ) AS Sold60,
  Format(
    Nz([AmtSold60], 0),
    "$#,###"
  ) AS Sales60,
  CInt(
    Nz([UnitsSoldLast90], 0)
  ) AS Sold90,
  Format(
    Nz([AmtSold90], 0),
    "$#,###"
  ) AS Sales90,
  Events_ByType_Last30.LastCall,
  Events_ByType_Last30.LastEmail,
  Events_ByType_Last30.LastVisit,
  Events_ByType_Last30.NextVisit,
  Events_ByType_Last30.EmailsPending,
  Events_ByType_Last30.CallsPending
FROM
  (
    (
      TBL_Employees
      LEFT JOIN BuyersProposalstatus_Last30 ON TBL_Employees.ID = BuyersProposalstatus_Last30.SalesPersonID
    )
    LEFT JOIN Events_ByType_Last30 ON TBL_Employees.ID = Events_ByType_Last30.SalesPersonID
  )
  LEFT JOIN RMKTUnitsSold30_60_90 ON TBL_Employees.EmpInt = RMKTUnitsSold30_60_90.RmkSm
WHERE
  (
    (
      (TBL_Employees.AccessGroup) Like "*RMKT*"
      Or (TBL_Employees.AccessGroup) Like "*ALL*"
    )
  )
ORDER BY
  TBL_Employees.[Employee Name];

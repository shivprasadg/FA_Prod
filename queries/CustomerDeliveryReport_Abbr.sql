SELECT
  vw_CustDelvReportSummed.Client,
  vw_CustDelvReportSummed.schgrp AS [Sch|Grp],
  IIf(
    [units] <> [UnitsInGroup],
    CStr([Units])& " (" & [UnitsInGroup] & ")",
    [UnitsInGroup]
  ) AS [#],
  vw_CustDelvReportSummed.UnitDescAbbr AS [Unit Description],
  vw_CustDelvReportSummed.Vendor,
  Format(
    [vw_CustDelvReportSummed].[Build],
    "m/d/yy"
  ) AS Build,
  vw_CustDelvReportSummed.Built,
  vw_CustDelvReportSummed.[Est Delv Dates],
  vw_CustDelvReportSummed.Delv,
  vw_CustDelvReportSummed.MSO,
  vw_CustDelvReportSummed.INV,
  vw_CustDelvReportSummed.Assignee,
  IIf(
    [CustPO1] = True,
    "C",
    IIf(
      IsDate([PO Date]),
      "Y",
      ""
    )
  ) AS PO,
  IIf(
    Nz([Interim bank], "")<> ""
    And [FundDate] > Date(),
    1,
    0
  ) AS IntDelv,
  vw_CustDelvReportSummed.[Interim Bank] AS IntBank
FROM
  vw_CustDelvReportSummed
ORDER BY
  IIf(
    Nz([Interim bank], "")<> ""
    And [FundDate] > Date(),
    1,
    0
  ),
  vw_CustDelvReportSummed.Client,
  vw_CustDelvReportSummed.schgrp;

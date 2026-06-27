SELECT
  TFR_2.FundingStatus,
  TFR_2.AccountName,
  TFR_2.FAID,
  TFR_2.Sch_Grp,
  TFR_2.UnitDesc,
  TFR_2.UnitCount,
  TFR_2.FundDate,
  TFR_2.Bank,
  Sum(TFR_2.OEC) AS TotalOEC,
  Sum(TFR_2.NetSyndFee) AS NetFee,
  Format(
    Sum([NetSyndFee])/ Sum([OEC]),
    "Percent"
  ) AS NetFeePct,
  Sum(
    [SynFeeProj] + [StripRent] + [TermPenaltyLPAF]
  ) AS ProjFee,
  Format(
    [ProjFee] / Sum([OEC]),
    "Percent"
  ) AS ProjFeePct,
  TFR_2.SGrpID
FROM
  TFR_2
GROUP BY
  TFR_2.FundingStatus,
  TFR_2.AccountName,
  TFR_2.FAID,
  TFR_2.Sch_Grp,
  TFR_2.UnitDesc,
  TFR_2.UnitCount,
  TFR_2.FundDate,
  TFR_2.Bank,
  TFR_2.SGrpID,
  CDbl(
    NumInString([Schedule])
  )& [UnitGroup]
ORDER BY
  TFR_2.FundingStatus DESC,
  TFR_2.AccountName,
  CDbl(
    NumInString([Schedule])
  )& [UnitGroup],
  TFR_2.FundDate DESC;

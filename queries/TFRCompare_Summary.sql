SELECT
  IIf(
    [Assigned_Indicator] = "F",
    "Syndicated",
    IIf(
      [Assigned_Indicator] = "P", "Forecast",
      "Portfolio"
    )
  ) AS FundStatus,
  TFRCompare.AccountName,
  Sum(TFRCompare.UnitCount) AS UnitCount,
  Sum(TFRCompare.GroupOEC) AS OEC,
  Sum(TFRCompare.SyndFee) AS SyndFee,
  Sum(TFRCompare.NetSyndFee) AS NetSyndFee,
  Sum(TFRCompare.NetSyndProj) AS NetSyndProj,
  TFRCompare.Bank AS TargetBank,
  Sum(TFRCompare.PDRent) AS PDRent,
  Sum(TFRCompare.PDRentProj) AS PDRentProj,
  Sum(TFRCompare.UTA) AS UTA,
  Sum(TFRCompare.UTAProj) AS UTAProj,
  Sum(TFRCompare.StripRent) AS StripRent,
  Sum(TFRCompare.TotalFee) AS TotalFee,
  Sum(TFRCompare.TotalFeeProj) AS TotalFeeProj,
  TFRCompare.clientID,
  TFRCompare.Assigned_Indicator,
  Sum(TFRCompare.PDRent) AS TotPDRent,
  Avg(TFRCompare.PDDays) AS AvgPDDays,
  Sum(TFRCompare.PmtAmt) AS RentIncome,
  Sum(TFRCompare.SyndFeePlus) AS SyndFeePlus
FROM
  TFRCompare
GROUP BY
  IIf(
    [Assigned_Indicator] = "F",
    "Syndicated",
    IIf(
      [Assigned_Indicator] = "P", "Forecast",
      "Portfolio"
    )
  ),
  TFRCompare.AccountName,
  TFRCompare.Bank,
  TFRCompare.clientID,
  TFRCompare.Assigned_Indicator
ORDER BY
  IIf(
    [Assigned_Indicator] = "F",
    "Syndicated",
    IIf(
      [Assigned_Indicator] = "P", "Forecast",
      "Portfolio"
    )
  ) DESC,
  TFRCompare.AccountName;

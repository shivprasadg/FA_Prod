SELECT
  vw_Syndication_Validate_Final.[Fund Status],
  vw_Syndication_Validate_Final.[Unit Status],
  vw_Syndication_Validate_Final.CompanyName,
  [MLno] & " - " & [Schedule] & " - " & [UnitGroup] AS [MLNo Sch Grp],
  vw_Syndication_Validate_Final.[FA Year],
  vw_Syndication_Validate_Final.UnitCount,
  vw_Syndication_Validate_Final.BLCDProj AS [Proj BLCD],
  vw_Syndication_Validate_Final.Assignee,
  vw_Syndication_Validate_Final.AssetID,
  vw_Syndication_Validate_Final.MYear,
  vw_Syndication_Validate_Final.Make,
  vw_Syndication_Validate_Final.Type,
  vw_Syndication_Validate_Final.[Unit Cost],
  vw_Syndication_Validate_Final.[Fund Date (Synd)],
  vw_Syndication_Validate_Final.[Synd_Fee (Avg)],
  vw_Syndication_Validate_Final.[Synd_Fee (Pct)],
  vw_Syndication_Validate_Final.[Strip Rent (Avg)],
  vw_Syndication_Validate_Final.[Strip Rent (Pct)],
  vw_Syndication_Validate_Final.[Term-Penalty (Avg)],
  vw_Syndication_Validate_Final.[Term-Penalty (Pct)],
  vw_Syndication_Validate_Final.[Per Diem],
  vw_Syndication_Validate_Final.[Per Diem (Pct)],
  vw_Syndication_Validate_Final.[UTA at Synd],
  vw_Syndication_Validate_Final.[UTA at Synd (Pct)],
  IIf(
    [PVFeesYN] = False,
    0,
    PV(
      ([PVRate] / 12),
      [FeeTerm],
      [FMSS],
      0,
      [PVType]
    )*-1
  ) AS [PV at SYND],
  IIf(
    [PV at SYND] > 0, [PV at SYND] / [Unit Cost],
    0
  ) AS [PV (Pct)],
  IIf(
    [IntrimLoanDays] > 0,
    [UnitIFValue] + FV(
      ([InterimLoanRate] / 360),
      [IntrimLoanDays],
      0,
      [UnitIFValue],
      1
    ),
    0
  ) AS [Interim Expense],
  IIf(
    [Interim Expense] <> 0, [Interim Expense] / [Unit Cost],
    0
  ) AS [Interim (Pct)],
  vw_Syndication_Validate_Final.Misc_1,
  vw_Syndication_Validate_Final.[Misc_1 (Pct)],
  vw_Syndication_Validate_Final.Misc_2,
  vw_Syndication_Validate_Final.[Misc_2 (Pct)],
  [Syndication_Total] - Abs([Interim Expense])+ [PV at SYND] AS [Syndication Total],
  (
    [Syndication_Total] - Abs([Interim Expense])+ [PV at SYND]
  )/ [Unit Cost] AS [Synd Total (Pct)],
  vw_Syndication_Validate_Final.[Fund Date (LO)],
  vw_Syndication_Validate_Final.[Proj Fee (Avg)],
  vw_Syndication_Validate_Final.[Proj Fee (Pct)],
  vw_Syndication_Validate_Final.[Proj Term-Penalty (Avg)],
  vw_Syndication_Validate_Final.[Proj Term-Penalty (Pct)],
  vw_Syndication_Validate_Final.[Proj Per Diem (Avg)],
  vw_Syndication_Validate_Final.[Proj Per Diem (Pct)],
  vw_Syndication_Validate_Final.[Proj UTA],
  vw_Syndication_Validate_Final.[Proj UTA (Pct)],
  PV(
    ([PVRate] / 12),
    [FeeTerm],
    [Proj FMSS],
    0,
    [PVType]
  )*-1 AS [Proj PV],
  IIf(
    [Proj PV] > 0, [Proj PV] / [Unit Cost],
    0
  ) AS [Proj PV (Pct)],
  [Projected_Total] + [Proj PV] AS [Projected Total],
  ([Projected_Total] + [Proj PV])/ [Unit Cost] AS [Proj Total (Pct)],
  vw_Syndication_Validate_Final.LeaseType,
  vw_Syndication_Validate_Final.PVType,
  vw_Syndication_Validate_Final.PVRate,
  vw_Syndication_Validate_Final.InterimLoanRate AS [Interim Loan Rate],
  vw_Syndication_Validate_Final.FeeTerm,
  vw_Syndication_Validate_Final.SyndNote,
  vw_Syndication_Validate_Final.SchRent,
  vw_Syndication_Validate_Final.SyndRent
FROM
  vw_Syndication_Validate_Final
ORDER BY
  vw_Syndication_Validate_Final.CompanyName;

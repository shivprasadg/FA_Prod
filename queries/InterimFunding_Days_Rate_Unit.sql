SELECT
  InterimFunding_Days_Rate.UnitID,
  InterimFunding_Days_Rate.GroupREF,
  InterimFunding_Days_Rate.InterimAmount,
  InterimFunding_Days_Rate.InterimIntRate,
  InterimFunding_Days_Rate.InterimLoanDays,
  IIf(
    Nz([InterimLoanDays], 0)< 1,
    0,
    [InterimAmount] + FV(
      ([InterimIntRate] / 360),
      [InterimLoanDays],
      0,
      [InterimAmount],
      1
    )
  ) AS UnitInterimExp,
  InterimFunding_Days_Rate.UnitCost
FROM
  InterimFunding_Days_Rate;

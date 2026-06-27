SELECT
  Units.UnitID,
  Units.UnitStatus,
  CCur(
    Nz([UnitIFAmt], 0)
  ) AS InterimAmount,
  Units.UnitIFAmt,
  SchGrp.SGrpSchID AS SchID,
  Units.GroupREF,
  UnitTotalOEC.UnitTotalCost,
  CDbl(
    Nz([InterimLoanRate], 0.025)
  ) AS InterimIntRate,
  IIf(
    IsDate([UnitIFDate])
    And IsDate([FundDate]),
    DateDiff("d", [UnitIFDate], [FundDate]),
    0
  ) AS InterimLoanDays,
  SchGrp.FundDate,
  SchGrp.LoanDate,
  Units.UnitIFDate,
  UnitTotalOEC.UnitCost
FROM
  (
    Units
    INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
  )
  INNER JOIN UnitTotalOEC ON Units.UnitID = UnitTotalOEC.AssetID;

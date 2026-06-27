SELECT
  FARentBillAssignedQ.LocationName,
  FARentBillAssignedQ.SchNo,
  FARentBillAssignedQ.SGrpNO AS GroupNo,
  FARentBillAssignedQ.LState,
  FARentBillAssignedQ.UnitTaxExempt,
  FARentBillAssignedQ.StTaxType,
  FARentBillAssignedQ.LTaxRate,
  Count(FARentBillAssignedQ.Unitvin) AS UnitCount,
  Sum(
    FARentBillAssignedQ.UnitTotalRent
  ) AS LineTot,
  Sum(FARentBillAssignedQ.FSLFee) AS FSLFee,
  Sum(FARentBillAssignedQ.RentNoTax) AS RentNoTax,
  Sum(FARentBillAssignedQ.TaxAmt) AS TaxAmt,
  Sum(
    FARentBillAssignedQ.UnitTotalRent
  ) AS UnitTotalRent,
  Sum(FARentBillAssignedQ.FSLTax) AS FSLTax
FROM
  FARentBillAssignedQ
GROUP BY
  FARentBillAssignedQ.LocationName,
  FARentBillAssignedQ.SchNo,
  FARentBillAssignedQ.SGrpNO,
  FARentBillAssignedQ.LState,
  FARentBillAssignedQ.UnitTaxExempt,
  FARentBillAssignedQ.StTaxType,
  FARentBillAssignedQ.LTaxRate;

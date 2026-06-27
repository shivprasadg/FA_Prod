UPDATE
  Units
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
SET
  Units.UnitTax = CCur(
    IIf(
      [UnitTaxExempt] = True,
      0,
      IIf(
        [UnitRent] > 0, [UnitRent], [UnitSRent]
      )* [LTaxRate]
    )
  )
WHERE
  (
    (
      (Units.UnitTax)< 1
      Or (Units.UnitTax) Is Null
    )
    AND (
      (Units.UnitStatus)<> "T"
    )
    AND (
      (Units.UnitTaxExempt)= False
    )
    AND (
      (Locations.LTaxRate)> 0
    )
    AND (
      (
        CCur(
          IIf(
            [UnitTaxExempt] = True,
            0,
            IIf(
              [UnitRent] > 0, [UnitRent], [UnitSRent]
            )* [LTaxRate]
          )
        )
      )> 0
    )
  );

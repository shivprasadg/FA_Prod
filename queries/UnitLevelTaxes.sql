SELECT
  Units.UnitID,
  Units.UnitStatus,
  Units.UnitInServ AS InServiceDate,
  Units.UnitPDStartDate AS PDStartDate,
  CCur(
    IIf(
      [UnitRent] > 0, [UnitRent], [UnitSRent]
    )
  ) AS Rent,
  CCur(
    Nz([SGrpFSLPmt], 0)
  ) AS FeeNoTax,
  CCur(
    Nz([SGrpRSLTPmt], 0)
  ) AS FeeTaxable,
  CCur(
    Round(
      IIf(
        [UnitTaxExempt] = True,
        0,
        ([UnitRent] + [FeeTaxable])* [LTaxRate]
      ),
      2
    )
  ) AS UnitTaxes,
  [Rent] + [FeeNoTax] + [FeeTaxable] + [UnitTaxes] AS RentPayment,
  Units.UnitTaxExempt,
  Locations.LTaxRate,
  Locations.Locationsn,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  TaxStatesBase.BaseTaxRate AS StateRate,
  TaxCountiesBase.BaseTaxRate AS CountyRate,
  TaxCitiesBase.BaseTaxRate AS CityRate
FROM
  (
    (
      (
        (
          Units
          LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
        )
        INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
      )
      LEFT JOIN TaxCitiesBase ON Locations.LCityID = TaxCitiesBase.TaxCityID
    )
    LEFT JOIN TaxCountiesBase ON Locations.LCountyID = TaxCountiesBase.TaxCountyID
  )
  LEFT JOIN TaxStatesBase ON Locations.LStateID = TaxStatesBase.TaxStatesID
WHERE
  (
    (
      (Units.UnitStatus)= "A"
      Or (Units.UnitStatus)= "B"
    )
    AND (
      (TaxCitiesBase.TaxTypeID)= 1
      Or (TaxCitiesBase.TaxTypeID) Is Null
    )
    AND (
      (TaxCountiesBase.TaxTypeId)= 1
      Or (TaxCountiesBase.TaxTypeId) Is Null
    )
    AND (
      (TaxStatesBase.TaxTypeId)= 1
      Or (TaxStatesBase.TaxTypeId) Is Null
    )
  )
ORDER BY
  Locations.Locationsn;

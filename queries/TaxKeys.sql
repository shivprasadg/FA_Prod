SELECT
  TaxStates.ISOCountryCode2,
  TaxStates.ID AS TaxStateID,
  TaxCounties.ID AS TaxCountyID,
  TaxCities.ID AS TaxCityID,
  TaxStates.PostalAbbreviation AS State,
  TaxCounties.CountyName AS County,
  TaxCities.CityName AS City
FROM
  (
    TaxStates
    INNER JOIN TaxCounties ON TaxStates.ID = TaxCounties.TaxStatesID
  )
  INNER JOIN TaxCities ON TaxCounties.ID = TaxCities.TaxCountyID;

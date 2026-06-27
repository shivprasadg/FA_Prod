SELECT
  vw_LocationDistribution_Taxes.Client,
  vw_LocationDistribution_Taxes.LocationName,
  vw_LocationDistribution_Taxes.Locationsn,
  vw_LocationDistribution_Taxes.LCity,
  vw_LocationDistribution_Taxes.LState,
  vw_LocationDistribution_Taxes.LCounty,
  vw_LocationDistribution_Taxes.LZip,
  Format(
    [vw_LocationDistribution_Taxes].[LUnits],
    "#,##0"
  ) AS LUnits,
  Format(
    [vw_LocationDistribution_Taxes].[FAUnits],
    "#,##0"
  ) AS FAUnits,
  Format(
    [vw_LocationDistribution_Taxes].[Distro],
    "Percent"
  ) AS Distro,
  Format(
    [vw_LocationDistribution_Taxes].[LTaxRate],
    "Percent"
  ) AS LTaxRate,
  vw_LocationDistribution_Taxes.LTaxOption,
  Format(
    [vw_LocationDistribution_Taxes].[LocTaxes],
    "Percent"
  ) AS LocTaxes,
  vw_LocationDistribution_Taxes.StateSales,
  vw_LocationDistribution_Taxes.CountySales,
  vw_LocationDistribution_Taxes.CitySales,
  vw_LocationDistribution_Taxes.StateRReceipt,
  vw_LocationDistribution_Taxes.LStateID,
  vw_LocationDistribution_Taxes.LCountyID,
  vw_LocationDistribution_Taxes.LCityID,
  vw_LocationDistribution_Taxes.LocationID,
  vw_LocationDistribution_Taxes.ClientGroupId,
  vw_LocationDistribution_Taxes.clientID
FROM
  vw_LocationDistribution_Taxes;

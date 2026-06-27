SELECT
  DISTINCT Units.UnitID,
  Unit_Vin_Remarket_Temp.Asset,
  Units.UnitUnitNum,
  Unit_Vin_Remarket_Temp.Unit,
  Unit_Vin_Remarket_Temp.VIN,
  Units.UnitVIN,
  Unit_Vin_Remarket_Temp.Make,
  Unit_Vin_Remarket_Temp.State,
  IIf(
    Trim([VIN])= Trim([UnitVIN]),
    "Yes",
    "No"
  ) AS [Match]
FROM
  Units
  INNER JOIN Unit_Vin_Remarket_Temp ON Units.UnitUnitNum = Unit_Vin_Remarket_Temp.Unit;

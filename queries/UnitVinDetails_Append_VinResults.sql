INSERT INTO UnitVinDetails (
  ClientID, UnitID, vin, modelyear, make,
  model, trimlevel, enginetype, transmission,
  drivelinetype, bodytype, vehicletype,
  vehicleclass, assyplant, cabtype,
  trailertype, bustype, fueltype, carlinecode,
  chassis, modelchange, country, motorcycletypes,
  horsepower, engineseriescode, gvwrclass,
  gvwrprodrange, wheelbase, axleconfiguration,
  manufacturer, enginemanufacturer,
  wmiidcode, plantcode, restraintsystem,
  series, enginemodifier, steering,
  emissionscode, brakesystem, frametype,
  hybridtype, enginecode, status, description
)
SELECT
  UnitVINResults.ClientID,
  Units.UnitID,
  UnitVINResults.vin,
  UnitVINResults.modelyear,
  UnitVINResults.make,
  UnitVINResults.model,
  UnitVINResults.trimlevel,
  UnitVINResults.enginetype,
  UnitVINResults.transmission,
  UnitVINResults.drivelinetype,
  UnitVINResults.bodytype,
  UnitVINResults.vehicletype,
  UnitVINResults.vehicleclass,
  UnitVINResults.assyplant,
  UnitVINResults.cabtype,
  UnitVINResults.trailertype,
  UnitVINResults.bustype,
  UnitVINResults.fueltype,
  UnitVINResults.carlinecode,
  UnitVINResults.chassis,
  UnitVINResults.modelchange,
  UnitVINResults.country,
  UnitVINResults.motorcycletypes,
  UnitVINResults.horsepower,
  UnitVINResults.engineseriescode,
  UnitVINResults.gvwrclass,
  UnitVINResults.gvwrprodrange,
  UnitVINResults.wheelbase,
  UnitVINResults.axleconfiguration,
  UnitVINResults.manufacturer,
  UnitVINResults.enginemanufacturer,
  UnitVINResults.wmiidcode,
  UnitVINResults.plantcode,
  UnitVINResults.restraintsystem,
  UnitVINResults.series,
  UnitVINResults.enginemodifier,
  UnitVINResults.steering,
  UnitVINResults.emissionscode,
  UnitVINResults.brakesystem,
  UnitVINResults.frametype,
  UnitVINResults.hybridtype,
  UnitVINResults.enginecode,
  UnitVINResults.status,
  UnitVINResults.description
FROM
  Units
  INNER JOIN (
    UnitVINResults
    LEFT JOIN UnitVinDetails ON UnitVINResults.vin = UnitVinDetails.vin
  ) ON Units.UnitVIN = UnitVINResults.VIN
WHERE
  (
    (
      (UnitVINResults.ClientID)= [Ask_ClientID]
    )
    AND (
      (UnitVINResults.make)<> ""
    )
    AND (
      (UnitVinDetails.vin) Is Null
    )
  );

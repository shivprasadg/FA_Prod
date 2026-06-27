SELECT
  FleetStudyNewVins.VINLenght,
  FleetStudyNewVins.UnitNumber,
  FleetStudyNewVins.OBCUnitNumber,
  FleetStudyNewVins.VIN,
  FleetStudyNewVins.UnitYear,
  FleetStudyNewVins.UnitMake,
  Nz(
    [FleetStudyNewVins].[UnitModel],
    [model]
  ) AS UnitModel,
  FleetStudyNewVins.Type,
  FleetStudyNewVins.SubType,
  FleetStudyNewVins.InServiceDate,
  IIf(
    [drivelinetype] Like '*X*', 'T/A',
    [FleetStudyNewVins].[Axle]
  ) AS Axle,
  UnitVINResults.drivelinetype AS AxleConfig,
  FleetStudyNewVins.UnitRent,
  FleetStudyNewVins.LeaseStartDate,
  FleetStudyNewVins.LeaseEndDate,
  FleetStudyNewVins.Unitcost,
  FleetStudyNewVins.[LeaseTerm ] AS Expr2,
  FleetStudyNewVins.Odometer,
  FleetStudyNewVins.OdometerDate,
  FleetStudyNewVins.Ownership,
  FleetStudyNewVins.Location,
  FleetStudyNewVins.State,
  FleetStudyNewVins.Plate,
  FleetStudyNewVins.ReeferHours,
  UnitVINResults.model,
  FleetImporter_Locations.LocationID AS LocID,
  FleetImporter_Locations.ClientREF,
  FleetStudyNewVins.Division,
  FleetStudyNewVins.DivisionID,
  FleetStudyNewVins.Engine,
  FleetStudyNewVins.Trans
FROM
  (
    (
      FleetStudyNewVins
      LEFT JOIN FleetImporter_ExisitingUnitNums ON FleetStudyNewVins.UnitNumber = FleetImporter_ExisitingUnitNums.UnitUnitNum
    )
    LEFT JOIN UnitVINResults ON FleetStudyNewVins.VIN = UnitVINResults.VIN
  )
  LEFT JOIN FleetImporter_Locations ON FleetStudyNewVins.Location = FleetImporter_Locations.LocationName
WHERE
  (
    (
      (
        FleetImporter_ExisitingUnitNums.UnitUnitNum
      ) Is Null
    )
  );

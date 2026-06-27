SELECT
  FleetStudyNewVins.VINLenght,
  FleetStudyNewVins.UnitNumber,
  FleetStudyNewVins.OBCUnitNumber,
  FleetStudyNewVins.VIN,
  FleetStudyNewVins.UnitYear,
  FleetStudyNewVins.UnitMake,
  FleetStudyNewVins.UnitModel AS UnitModel,
  FleetStudyNewVins.Type,
  FleetStudyNewVins.SubType,
  FleetStudyNewVins.InServiceDate,
  Nz(
    [FleetStudyNewVins].[Axle], "TBD"
  ) AS Axle,
  FleetStudyNewVins.UnitRent,
  FleetStudyNewVins.LeaseStartDate,
  FleetStudyNewVins.LeaseEndDate,
  FleetStudyNewVins.Unitcost,
  FleetStudyNewVins.[LeaseTerm ],
  FleetStudyNewVins.Odometer,
  FleetStudyNewVins.OdometerDate,
  FleetStudyNewVins.Location,
  FleetStudyNewVins.State,
  FleetStudyNewVins.Plate,
  FleetStudyNewVins.ReeferHours,
  FleetImporter_Locations.LocationID AS LocID,
  FleetImporter_Locations.ClientREF
FROM
  (
    FleetStudyNewVins
    LEFT JOIN FleetImporter_ExisitingUnitNums ON FleetStudyNewVins.UnitNumber = FleetImporter_ExisitingUnitNums.UnitUnitNum
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

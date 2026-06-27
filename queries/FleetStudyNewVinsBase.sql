SELECT
  IIf(
    Len(
      Trim([VIN])
    )< 17,
    "Short",
    IIf(
      Len(
        Trim([VIN])
      )> 17,
      "Long",
      "Good"
    )
  ) AS VINLenght,
  FleetStudyImportBase.UnitNumber,
  FleetStudyImportBase.OBCUnitNumber,
  Trim([FleetStudyImportBase].[VIN]) AS VIN,
  FleetStudyImportBase.UnitYear,
  FleetStudyImportBase.UnitMake,
  FleetStudyImportBase.UnitModel,
  FleetStudyImportBase.Type,
  FleetStudyImportBase.SubType,
  FleetStudyImportBase.InServiceDate,
  FleetStudyImportBase.Axle,
  FleetStudyImportBase.UnitRent,
  FleetStudyImportBase.LeaseStartDate,
  FleetStudyImportBase.LeaseEndDate,
  FleetStudyImportBase.Unitcost,
  FleetStudyImportBase.Division,
  FleetStudyImportBase.[LeaseTerm ],
  FleetStudyImportBase.[Odometer ],
  FleetStudyImportBase.OdometerDate,
  FleetStudyImportBase.Location,
  FleetStudyImportBase.State,
  FleetStudyImportBase.Plate,
  FleetStudyImportBase.ReeferHours
FROM
  FleetStudyImportBase
  LEFT JOIN Units ON FleetStudyImportBase.VIN = Units.UnitVIN
WHERE
  (
    (
      (Units.UnitVIN) Is Null
    )
  );

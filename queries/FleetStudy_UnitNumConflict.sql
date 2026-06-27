SELECT
  FleetStudyNewVinsBase.VINLenght,
  UnitLastOBCReport.LogActivityDate,
  UnitVinDecoder.status AS DecodeStatus,
  FleetStudyNewVinsBase.UnitNumber,
  FleetStudyNewVinsBase.OBCUnitNumber,
  FleetStudyNewVinsBase.VIN,
  FleetStudyNewVinsBase.UnitYear,
  FleetStudyNewVinsBase.UnitMake,
  FleetStudyNewVinsBase.UnitModel,
  FleetStudyNewVinsBase.Type,
  FleetStudyNewVinsBase.SubType,
  FleetStudyNewVinsBase.InServiceDate,
  FleetStudyNewVinsBase.Axle,
  FleetStudyNewVinsBase.UnitRent,
  FleetStudyNewVinsBase.LeaseStartDate,
  FleetStudyNewVinsBase.LeaseEndDate,
  FleetStudyNewVinsBase.Unitcost,
  FleetStudyNewVinsBase.Division,
  FleetStudyNewVinsBase.[LeaseTerm ],
  FleetStudyNewVinsBase.[Odometer ],
  FleetStudyNewVinsBase.OdometerDate,
  FleetStudyNewVinsBase.Location,
  FleetStudyNewVinsBase.State,
  FleetStudyNewVinsBase.Plate,
  FleetStudyNewVinsBase.ReeferHours,
  FleetImporter_ExisitingUnitNums.UnitUnitNum,
  vw_SixKeys.UnitDescShortVer,
  vw_SixKeys.VIN
FROM
  (
    (
      (
        FleetStudyNewVinsBase
        LEFT JOIN FleetImporter_ExisitingUnitNums ON FleetStudyNewVinsBase.UnitNumber = FleetImporter_ExisitingUnitNums.UnitUnitNum
      )
      LEFT JOIN vw_SixKeys ON FleetImporter_ExisitingUnitNums.UnitID = vw_SixKeys.AssetID
    )
    LEFT JOIN UnitLastOBCReport ON vw_SixKeys.AssetID = UnitLastOBCReport.UnitID
  )
  LEFT JOIN UnitVinDecoder ON vw_SixKeys.AssetID = UnitVinDecoder.UnitID
WHERE
  (
    (
      (
        FleetImporter_ExisitingUnitNums.UnitUnitNum
      ) Is Not Null
    )
  );

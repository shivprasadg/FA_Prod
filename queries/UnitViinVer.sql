SELECT
  Units.*,
  RmktVinDecodeT.[Engine Manufacturer],
  RmktVinDecodeT.[Engine Series],
  RmktVinDecodeT.[Engine Series],
  RmktVinDecodeT.[Engine Configuration],
  RmktVinDecodeT.[Engine Capacity],
  RmktVinDecodeT.[Engine Capacity],
  RmktVinDecodeT.[Fuel Type],
  RmktVinDecodeT.[Fuel Type]
FROM
  Units
  LEFT JOIN RmktVinDecodeT ON Units.UnitID = RmktVinDecodeT.UnitID;

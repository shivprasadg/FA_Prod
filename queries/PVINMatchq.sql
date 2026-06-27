SELECT
  PMatchOne.[VIN #],
  Units.UnitVIN,
  PMatchOne.[License Plate State]
FROM
  Units
  INNER JOIN PMatchOne ON Units.UnitVIN = PMatchOne.[VIN #];

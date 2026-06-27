UPDATE
  Units
  INNER JOIN upload ON Units.UnitID = upload.[FA#]
SET
  Units.unitunitnum = [costco unit #],
  Units.unitvin = [vin #],
  Units.unitbuilddate = [build date];

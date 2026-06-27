SELECT
  Units.unitunitnum,
  Units.unitvin,
  Units.UnitID,
  ChildParts.UnitREF,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildVIN,
  ChildParts.ChildVendorREF,
  ChildParts.ChildEXCofa,
  ChildParts.ChildVIN,
  ChildParts.REFENG
FROM
  (
    Units
    INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
  )
  INNER JOIN Wmart2 ON Units.UnitVIN = Wmart2.unitvin;

UPDATE
  Units
SET
  Units.ObcUnitNumber = Switch(
    IsNull([units].[UnitUnitNum])
    And [UnitStatus] Like "*S*",
    "Sold" & Right([UnitVin], 6),
    IsNull([units].[UnitUnitNum])
    And [UnitStatus] = "O",
    "OL" & Right([UnitVin], 6),
    IsNull([units].[UnitUnitNum])
    And [UnitStatus] = "T",
    "TRK" & Right([UnitVin], 6),
    IsNull([units].[UnitUnitNum])
    And [UnitStatus] = "A",
    "ATL" & Right([UnitVin], 6),
    Not IsNull([units].[UnitUnitNum]),
    [UnitUnitNum]
  )
WHERE
  (
    (
      (Units.ObcUnitNumber) Is Null
    )
    AND (
      (Units.UnitStatus)<> "B"
    )
  );

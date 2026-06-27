UPDATE
  Units
SET
  Units.UnitClassType = IIf(
    "UnitSubType" = "Tractor", "Tractor",
    "Trailer"
  )
WHERE
  (
    (
      ([Units].[UnitClassType]) Is Null
    )
  );

UPDATE
  Units
SET
  Units.GroupREF = [forms]![moveunit]![newgroup]
WHERE
  (
    (
      (Units.UnitID)= [forms]![moveunit]![unit]
    )
  );

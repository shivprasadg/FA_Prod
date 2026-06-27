SELECT
  UnitNotes.*
FROM
  UnitNotes
WHERE
  (
    (
      (UnitNotes.UnitRef)= [forms]![UnitsForm]![unitid]
    )
  );

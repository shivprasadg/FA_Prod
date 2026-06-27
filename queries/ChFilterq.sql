SELECT
  ChildParts.*
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.UnitREF)= [forms]![UnitsForm]![unitid]
    )
  );

SELECT
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildVIN,
  ChildParts.REFENG,
  ChildParts.COMPRESSS,
  ChildParts.COMPRESSS,
  ChildParts.MICROPROC,
  ChildParts.ELECDISP,
  ChildParts.[ARB ID]
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType)<> "decal"
      And (ChildParts.ChildType)<> "freight"
      And (ChildParts.ChildType)<> "warranty"
    )
  );

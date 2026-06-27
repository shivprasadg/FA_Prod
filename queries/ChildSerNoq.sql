SELECT
  ChildParts.UnitREF,
  ChildParts.ChildYear,
  ChildParts.ChildVIN AS [CPSerial No],
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.REFENG,
  ChildParts.COMPRESSS,
  ChildParts.MICROPROC,
  ChildParts.ELECDISP,
  ChildParts.[ARB ID]
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Not Like "decal*"
      And (ChildParts.ChildType) Not Like "*warranty"
      And (ChildParts.ChildType) Not Like "*Installation*"
      And (ChildParts.ChildType)<> "freight"
      And (ChildParts.ChildType) Not Like "*modification*"
    )
  );

SELECT
  ChildParts.UnitREF,
  ChildParts.ChildVIN,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildEXCofa,
  IIf(
    IsNull([ChildVIN]),
    Null,
    Trim([childyear])& "   " & Trim([childmake])& IIf(
      [ChildClass] = ".",
      Null,
      "   " & Trim([ChildClass])
    )& "   " & Trim([childtype])& "    " & Trim([ChildVIN])
  ) AS ChildDesc
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildEXCofa)= No
    )
  )
ORDER BY
  ChildParts.UnitREF,
  ChildParts.ChildVIN;

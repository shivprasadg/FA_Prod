INSERT INTO ChildParts (
  UnitREF, ChildYear, ChildMake, ChildClass,
  ChildType, childFinalCost, ChildVendorREF,
  ChildStatus, ChildEXCofa, PartAddedDate
)
SELECT
  Units.UnitID,
  [forms]![add additional Child]![AddCYR] AS YR,
  [forms]![add additional Child]![AddCMake] AS Make,
  [forms]![add additional Child]![AddcModel] AS Model,
  [forms]![add additional Child]![AddCType] AS Type,
  [forms]![add additional Child]![AddCCost] AS Cost,
  [forms]![add additional Child]![AddCvendor] AS Vendor,
  [forms]![add additional Child]![AddCstatus] AS Status,
  [forms]![add additional Child]![AccCCofa] AS OnCofA,
  Date() AS CHAdd
FROM
  Units
WHERE
  (
    (
      (Units.GroupREF)= [forms]![add additional child]![selgrp]
    )
  )
ORDER BY
  Units.UnitID;

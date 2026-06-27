SELECT
  ChildParts.UnitREF,
  ChildParts.ChildVIN,
  UCase(
    [childyear] & " " & [childmake] & " " & [ChildClass] & " " & [childtype]
  ) AS Childdesc,
  ChildParts.ChildInvoiceNum,
  ChildParts.childFinalCost,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ExclFrGrpOEC,
  ChildParts.ChildEXCofa,
  ChildParts.ChildDeposit
FROM
  ChildParts
ORDER BY
  ChildParts.UnitREF;

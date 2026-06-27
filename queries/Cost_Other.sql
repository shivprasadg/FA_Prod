SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Other,
  [ChildMake] & " - " & [ChildType] AS OtherMake,
  ChildParts.ChildInvoiceNum AS OtherInvoice,
  ChildParts.ChildVendorPytDate AS OtherPaid,
  ChildParts.DrawDownId AS OtherDraw,
  ChildParts.ChildType
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  [ChildMake] & " - " & [ChildType],
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId,
  ChildParts.ChildType
HAVING
  (
    (
      (ChildParts.ChildInvoiceNum) Is Not Null
    )
    AND (
      (ChildParts.ChildType) Not Like "*Refrig*"
      And (ChildParts.ChildType) Not Like "*Decal*"
      And (ChildParts.ChildType) Not Like "*Freight*"
      And (ChildParts.ChildType) Not Like "*OBC*"
      And (ChildParts.ChildType) Not Like "*XATA*"
      And (ChildParts.ChildType) Not Like "*Warranty*"
      And (ChildParts.ChildType) Not Like "*Body*"
      And (ChildParts.ChildType) Not Like "*Tank*"
      And (ChildParts.ChildType) Not Like "*Tax*"
      And (ChildParts.ChildType) Not Like "*Inspection*"
      And (ChildParts.ChildType) Not Like "*PTO*"
      And (ChildParts.ChildType) Not Like "*Oil*"
      And (ChildParts.ChildType) Not Like "*Cable*"
    )
  );

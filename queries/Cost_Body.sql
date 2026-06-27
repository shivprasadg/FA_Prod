SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Body
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Body*"
      Or (ChildParts.ChildType) Like "*Tank*"
    )
  )
GROUP BY
  ChildParts.UnitREF;

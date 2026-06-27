SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Decals,
  ChildParts.DrawDownId AS DecalsDraw
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.DrawDownId;

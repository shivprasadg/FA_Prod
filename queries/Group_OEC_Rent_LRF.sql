SELECT
  Units.GroupREF AS UnitGroupID,
  Sum(Units.UnitFinalCost) AS UnitCost,
  Sum(
    CCur(
      Nz([ChildPartCost], 0)
    )
  ) AS ChildCost,
  Sum(
    CCur(
      [UnitFinalCost] + Nz([ChildPartCost], 0)
    )
  ) AS GroupOEC,
  CCur(
    Sum(
      Nz([UnitRent], 0)
    )
  ) AS GroupRent,
  IIf(
    [GroupRent] = 0
    Or [GroupOEC] = 0,
    0,
    CDbl([GroupRent] / [GroupOEC])
  ) AS GroupLRF,
  CCur(
    Sum(
      Nz([UnitSRent], 0)
    )
  ) AS GroupSalesRent,
  IIf(
    [GroupSalesRent] = 0
    Or [GroupOEC] = 0,
    0,
    CDbl([GroupSalesRent] / [GroupOEC])
  ) AS GroupLoLRF,
  IIf(
    Sum(
      Nz([FirstExtRent], 0)
    )= 0,
    0,
    CCur(
      Sum([FirstExtRent])/ Nz([UnitsInGroup], 1)
    )
  ) AS GroupExtRent,
  IIf(
    Sum(
      Nz([Units.MTMRent])
    )= 0,
    0,
    CCur(
      Sum([Units.MTMRent])/ Nz([UnitsInGroup], 1)
    )
  ) AS GroupMTMRent,
  SchGrp.UnitsInGroup,
  IIf(
    [GroupMTMRent] = 0
    Or [GroupOEC] = 0,
    0,
    CDbl(
      [GroupMTMRent] /([GroupOEC] / [Unitsingroup])
    )
  ) AS MTMLRF,
  IIf(
    [GroupExtRent] = 0
    Or [GroupOEC] = 0,
    0,
    CDbl(
      [GroupExtRent] /([GroupOEC] / [UnitsInGroup])
    )
  ) AS EXTLRF,
  Sum(Units.UnitSalesResid) AS GroupSalesRV,
  Sum(Units.UnitSynResid) AS GroupSyndRV,
  Sum(
    [VendorCost] + Nz([PartsCost], 0)
  ) AS VendorPlus
FROM
  (
    Units
    LEFT JOIN ChildPart_Cost_OEC_summed ON Units.UnitID = ChildPart_Cost_OEC_summed.UnitId
  )
  INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
GROUP BY
  Units.GroupREF,
  SchGrp.UnitsInGroup;

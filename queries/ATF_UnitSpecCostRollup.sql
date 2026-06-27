SELECT
  ATFUnitSpec.ATFID,
  ATFUnitSpec.ATFGroupSpecID,
  Sum(ATFUnitSpec.NumUnitsInSpec) AS NumUnitsInSpec,
  Sum(ATFUnitSpec.UnitCost) AS UnitCost,
  Sum(
    CCur(
      Nz([ChildPartCost], 0)
    )
  ) AS PartCost,
  [UnitCost] + CCur(
    Nz(
      Sum([ChildPartCost]),
      0
    )
  ) AS TotalUnitCost,
  Sum(ATFUnitSpec.SchRent) AS SchRent,
  IIf(
    Nz([SchRent], 0)= 0
    Or Nz([TotalUnitCost], 0)= 0,
    Null,
    Round([SchRent] / [TotalUnitCost], 5)
  ) AS LRFSch,
  Sum(ATFUnitSpec.FirstExtendedRent) AS FirstExtendedRent,
  IIf(
    Nz([FirstExtendedRent], 0)= 0
    Or Nz([TotalUnitCost], 0)= 0,
    Null,
    Round(
      [FirstExtendedRent] / [TotalUnitCost],
      5
    )
  ) AS LRFExt,
  Sum(ATFUnitSpec.MTMRent) AS MTMRent,
  IIf(
    Nz([MTMRent], 0)= 0
    Or Nz([TotalUnitCost], 0)= 0,
    Null,
    Round([MTMRent] / [TotalUnitCost], 5)
  ) AS LRFMTM,
  Sum(ATFUnitSpec.PerDiemRent) AS PerDiemRent,
  Avg(ATFUnitSpec.FirstExtendedTerm) AS FirstExtendedTerm,
  Sum(ATFUnitSpec.SalesResidual) AS SalesResidual,
  Count(ATFUnitSpec.ATFUnitSpecID) AS SpecsIn,
  Sum(
    CCur(
      Nz([ATFUnitSpec].[UTARebate], 0)
    )
  ) AS UTARebate
FROM
  ATFUnitSpec
  LEFT JOIN ATF_ChildPartCost ON ATFUnitSpec.ATFUnitSpecID = ATF_ChildPartCost.ATFUnitSpecID
GROUP BY
  ATFUnitSpec.ATFID,
  ATFUnitSpec.ATFGroupSpecID;

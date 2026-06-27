SELECT
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Unitnum AS UnitNumber,
  vw_SixKeys.VIN,
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  [Axle] & ", " & [UnitSubtype] & " " & [UnitType] AS Type,
  CCur([UnitCost]) AS [Unit Cost],
  CCur(
    Nz([ChildPartCost], 0)
  ) AS [Child Cost],
  [locations].[LCity] & " " & [locations].[LState] & " - " & [locations].[LocationSN] AS [Domicile Location],
  vw_SixKeys.EsitmatedDeliveryDate AS [Target Delivery],
  "" AS [Spec Letter],
  vw_SixKeys.ActualDeliveryDate AS [Actual Delivery Date],
  IIf(
    [AllowedMilesPeriod] Like "*Term*",
    (
      (
        [SgprAllow1] /(
          [vw_SixKeys].[LeaseTermPrimary] - Nz([LeaseTermNotice], 0)
        )
      )* 12
    ),
    [SgprAllow1]
  ) AS [Annual Mileage Allowance],
  Units.unitsrent AS [Monthly Rent per Schedule Document],
  [vw_SixKeys].[LeaseTermPrimary] - Nz([LeaseTermNotice], 0) AS [Primary Lease Term],
  Units.unitrent AS [Final Monthly Rent per RAL Document],
  SchGrp.LeaseExtOption1Term AS [ExtendIT Term # Months],
  SchGrp.LeaseExtOption1Rent AS [ExtendIT Monthly Rent],
  SchGrp.MTMRent AS [M-T-M Rent],
  vw_SixKeys.SchID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID
FROM
  (
    (
      (
        vw_SixKeys
        INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
      )
      INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    LEFT JOIN ChildPart_Cost_OEC_summed ON Units.UnitID = ChildPart_Cost_OEC_summed.UnitId
  )
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
ORDER BY
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  [locations].[LCity] & " " & [locations].[LState] & " - " & [locations].[LocationSN];

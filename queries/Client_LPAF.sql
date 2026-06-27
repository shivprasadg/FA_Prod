SELECT
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Unitnum AS UnitNumber,
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
  IIf(
    [AllowedMilesPeriod] Like "*Term*",
    (
      (
        [SgprAllow1] /(
          [SchGrp].[LeaseTermPrimary] - Nz([SchGrp].[LeaseTermNotice], 0)
        )
      )* 12
    ),
    [SgprAllow1]
  ) AS [Annual Mileage],
  Units.unitsrent AS [LO Rent],
  [SchGrp].[LeaseTermPrimary] - Nz([SchGrp].[LeaseTermNotice]) AS [Primary Lease Term],
  Units.unitrent AS [Synd Rent],
  SchGrp.LeaseExtOption1Term AS [Lease Extend Term],
  SchGrp.LeaseExtOption1Rent AS [Lease Extend Rent],
  SchGrp.MTMRent,
  vw_SixKeys.SchID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.UnitID,
  vw_SixKeys.ClientGroupId
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
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  [locations].[LCity] & " " & [locations].[LState] & " - " & [locations].[LocationSN];

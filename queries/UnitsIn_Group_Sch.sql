SELECT
  Units.UnitID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.GroupREF,
  SchGrp.SGrpID,
  Units.UnitStatus,
  Units.UnitEstDelDate,
  Units.UnitVIN,
  Units.UnitUnitNum,
  Units.UnitType,
  Units.UnitSubType,
  (
    Right([UNITVIN], 6)
  ) AS vin4,
  lookup_groupTypes.GrpTypAbrev,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitCOASent,
  Units.UnitCOABack,
  Units.UnitAcceptDate,
  Units.UnitFinalCost,
  Sum(ChildParts.childFinalCost) AS SumOfchildFinalCost,
  Units.unitoffleasedt,
  Units.UnitActualDelDate,
  "(" & [lOCATIONSN] & ")" & " " & [LCity] & " " & [locations].[LState] AS ULoc,
  lookup_GroupMakes.MakeAbrev,
  Units.UnitDocAccpt,
  Units.unitinserv,
  Units.UnitInServInd,
  Units.UnitPlateNum,
  Units.UnitOLInd,
  Units.unitunitnum2,
  Units.Axle,
  Schedule.SchID,
  SchGrp.SGrpID,
  LookUp_UnitDesc.DescriptionLong AS UnitDesc,
  Nz(
    [units].[UnitCoABack], [Units].[UnitCofACopy]
  ) AS [CofABck-Cpy],
  Units.UnitCOABack,
  Units.UnitCofACopy
FROM
  (
    (
      (
        (
          (
            (
              Clients
              INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
            )
            INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
          )
          INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
        )
        INNER JOIN (
          (
            Units
            LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
          )
          LEFT JOIN lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake
        ) ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
    )
    LEFT JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
  )
  INNER JOIN LookUp_UnitDesc ON Units.UnitID = LookUp_UnitDesc.AssetID
GROUP BY
  Units.UnitID,
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.GroupREF,
  SchGrp.SGrpID,
  Units.UnitStatus,
  Units.UnitEstDelDate,
  Units.UnitVIN,
  Units.UnitUnitNum,
  Units.UnitType,
  Units.UnitSubType,
  (
    Right([UNITVIN], 6)
  ),
  lookup_groupTypes.GrpTypAbrev,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitCOASent,
  Units.UnitAcceptDate,
  Units.UnitFinalCost,
  Units.unitoffleasedt,
  Units.UnitActualDelDate,
  "(" & [lOCATIONSN] & ")" & " " & [LCity] & " " & [locations].[LState],
  lookup_GroupMakes.MakeAbrev,
  Units.UnitDocAccpt,
  Units.unitinserv,
  Units.UnitInServInd,
  Units.UnitPlateNum,
  Units.UnitOLInd,
  Units.unitunitnum2,
  Units.Axle,
  Schedule.SchID,
  SchGrp.SGrpID,
  LookUp_UnitDesc.DescriptionLong,
  Units.UnitCOABack,
  Units.UnitCofACopy,
  SchGrp.SGrpID,
  Units.UnitCOABack
HAVING
  (
    (
      (Units.GroupREF)= [Forms]![GroupForm]![GroupID]
    )
  )
ORDER BY
  Units.UnitID,
  Units.UnitType;

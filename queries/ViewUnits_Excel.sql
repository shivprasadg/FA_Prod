SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitStatus,
  Units.UnitEstDelDate,
  Units.UnitVIN,
  Units.UnitUnitNum,
  Units.UnitType,
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
  Units.UnitDocAccpt,
  Units.unitinserv,
  Units.UnitInServInd,
  Units.UnitPlateNum,
  Units.UnitID AS AssetID,
  Units.UnitOLInd,
  Units.unitunitnum2,
  Units.Axle
FROM
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
GROUP BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitStatus,
  Units.UnitEstDelDate,
  Units.UnitVIN,
  Units.UnitUnitNum,
  Units.UnitType,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitCOASent,
  Units.UnitCOABack,
  Units.UnitAcceptDate,
  Units.UnitFinalCost,
  Units.unitoffleasedt,
  Units.UnitActualDelDate,
  "(" & [lOCATIONSN] & ")" & " " & [LCity] & " " & [locations].[LState],
  Units.UnitDocAccpt,
  Units.unitinserv,
  Units.UnitInServInd,
  Units.UnitPlateNum,
  Units.UnitID,
  Units.UnitOLInd,
  Units.unitunitnum2,
  Units.Axle,
  Units.GroupREF,
  SchGrp.SGrpID,
  (
    Right([UNITVIN], 6)
  ),
  lookup_groupTypes.GrpTypAbrev,
  lookup_GroupMakes.MakeAbrev
HAVING
  (
    (
      (MstrLease.MLNo)= Forms![UnitsForm-Sub]!MLNo
    )
    And (
      (Schedule.SchNo)= Forms![UnitsForm-Sub]!SchNo
    )
    And (
      (SchGrp.SGrpNO)= Forms![UnitsForm-Sub]!SGrpNO
    )
  )
ORDER BY
  Units.UnitID,
  Units.UnitType;

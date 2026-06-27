SELECT
  Schedule.SchNo,
  Locations.LocationName,
  Units.unitunitnum,
  Units.unitvin,
  IIf(
    IsNull([unitinserv]),
    [UnitAcceptDate],
    [unitinserv]
  ) AS InServDt,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  SchGrp.BLCD,
  PMs.PMDate,
  PMs.PMno,
  PMs.MI,
  PMs.hr,
  PMs.[Work Type],
  PMs.[Work Desc]
FROM
  (
    (
      (
        MstrLease
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN (
      Units
      INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN PMs ON Units.UnitID = PMs.UnitRef
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Schedule.SchID,
  Locations.LocationName,
  Units.unitunitnum,
  Schedule.SchID,
  PMs.PMDate,
  PMs.PMno,
  PMs.[Work Desc];

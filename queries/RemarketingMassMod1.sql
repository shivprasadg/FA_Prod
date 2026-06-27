SELECT
  Units.UnitID,
  Units.Unitunitnum AS UnitNum,
  Units.Unitvin,
  Schedule.SchID,
  Schedule.SchNo,
  Units.UnitStatus,
  Rmkt.*,
  Units.unitoffleasedt,
  Units.LocationParkedID
FROM
  (
    Schedule
    INNER JOIN (
      (
        SchGrp
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN RMKT_MMod_Work ON Units.UnitID = RMKT_MMod_Work.UnitID
WHERE
  (
    (
      (Units.UnitStatus)<> "B"
    )
  );

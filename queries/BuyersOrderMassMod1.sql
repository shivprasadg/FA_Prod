SELECT
  Units.UnitID,
  Units.Unitunitnum AS UnitNum,
  Units.Unitvin,
  Schedule.SchID,
  Schedule.SchNo,
  Units.UnitStatus,
  Last,
  Rmkt.*,
  Units.unitoffleasedt,
  Units.LocationParkedID,
  Units.clientGroupId,
  SchGrp.AssigneeBankID
FROM
  (
    Schedule
    INNER JOIN (
      (
        (
          SchGrp
          INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
        )
        LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
      )
      LEFT JOIN (
        SELECT
          AssetID, LastDirection AS Last
        FROM
          RMKTUnitTitleLog_Work
      ) AS LT ON Units.UnitID = LT.AssetID
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN BuyersOrder_MMod_Work ON Units.UnitID = BuyersOrder_MMod_Work.UnitID
WHERE
  (
    (
      (Units.UnitStatus)<> "B"
    )
  );

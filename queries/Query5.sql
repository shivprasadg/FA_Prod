SELECT
  R.*,
  U.UnitID,
  U.UnitUnitNum AS UnitNum,
  U.UnitVIN,
  U.UnitStatus,
  U.UnitOffLeaseDt,
  U.LocationParkedID,
  U.ClientGroupID,
  SG.AssigneeBankID,
  S.SchID,
  S.SchNo,
  RMKTUnitTitleLog_Work.LastEntry
FROM
  (
    (
      (
        (
          RMKT AS R
          INNER JOIN BuyersOrder_MMod_Work AS BO ON R.UnitRef = BO.UnitID
        )
        INNER JOIN Units AS U ON R.UnitRef = U.UnitID
      )
      INNER JOIN SchGrp AS SG ON U.GroupREF = SG.SGrpID
    )
    INNER JOIN Schedule AS S ON SG.SGrpSchID = S.SchID
  )
  INNER JOIN RMKTUnitTitleLog_Work ON BO.UnitID = RMKTUnitTitleLog_Work.AssetID;

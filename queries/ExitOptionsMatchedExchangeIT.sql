SELECT
  UnitExitOption.UnitID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  UnitExitOption.OptionTypeID,
  UnitExitOption.ReplacedByScheduleID,
  UnitExitOption.ReplacedByGroupID,
  UnitExitOption.ReplacedByUnitID,
  Schedule.SchNo AS NewSchedule,
  SchGrp.SGrpNO AS NewGroup,
  Units.UnitID AS NewAssetID,
  Units.UnitUnitNum AS NewUnitNum
FROM
  (
    (
      (
        UnitExitOption
        INNER JOIN Schedule ON UnitExitOption.ReplacedByScheduleID = Schedule.SchID
      )
      INNER JOIN SchGrp ON UnitExitOption.ReplacedByGroupID = SchGrp.SGrpID
    )
    LEFT JOIN Units ON UnitExitOption.ReplacedByUnitID = Units.UnitID
  )
  INNER JOIN vw_SixKeys ON UnitExitOption.UnitID = vw_SixKeys.AssetID
GROUP BY
  UnitExitOption.UnitID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  UnitExitOption.OptionTypeID,
  UnitExitOption.ReplacedByScheduleID,
  UnitExitOption.ReplacedByGroupID,
  UnitExitOption.ReplacedByUnitID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum
HAVING
  (
    (
      (UnitExitOption.OptionTypeID)= 1
    )
  );

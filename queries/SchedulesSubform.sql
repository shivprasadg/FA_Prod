SELECT
  SchedulesSubform_Pre.*,
  ScheduleExchangedBy.NewSchGrp,
  ScheduleExchangedBy.ReplacedByScheduleID,
  ScheduleExchangedBy.ReplacedByGroupID
FROM
  SchedulesSubform_Pre
  LEFT JOIN ScheduleExchangedBy ON SchedulesSubform_Pre.SchID = ScheduleExchangedBy.SchID;

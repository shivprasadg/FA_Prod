SELECT
  ExitOptionsMatchedExchangeIT.SchID,
  Max([NewSchedule] & " | " & [NewGroup]) AS NewSchGrp,
  First(
    ExitOptionsMatchedExchangeIT.ReplacedByScheduleID
  ) AS ReplacedByScheduleID,
  First(
    ExitOptionsMatchedExchangeIT.ReplacedByGroupID
  ) AS ReplacedByGroupID
FROM
  ExitOptionsMatchedExchangeIT
GROUP BY
  ExitOptionsMatchedExchangeIT.SchID;

SELECT
  vw_ScheduleHealthCheck.*
FROM
  vw_ScheduleHealthCheck
WHERE
  (
    (
      (vw_ScheduleHealthCheck.FundX)> #12/31/2021#
    )
  )
ORDER BY
  vw_ScheduleHealthCheck.Lessee,
  vw_ScheduleHealthCheck.Schedule DESC;

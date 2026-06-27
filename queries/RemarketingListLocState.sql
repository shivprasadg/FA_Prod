SELECT
  RemarketingReportAllUnits.LState
FROM
  RemarketingReportAllUnits
GROUP BY
  RemarketingReportAllUnits.LState
HAVING
  (
    (
      (
        RemarketingReportAllUnits.LState
      ) Is Not Null
      Or (
        RemarketingReportAllUnits.LState
      )<> ''
    )
  )
ORDER BY
  RemarketingReportAllUnits.LState;

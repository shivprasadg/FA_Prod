SELECT
  Schedule.SchFAID AS FAID,
  Schedule.SchNo
FROM
  Schedule
GROUP BY
  Schedule.SchFAID,
  Schedule.SchNo
HAVING
  (
    (
      (Schedule.SchFAID)<> ""
    )
  );

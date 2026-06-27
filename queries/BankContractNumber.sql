SELECT
  Schedule.SchID,
  [SchBnkClNO] & "-" & [SgrpBnkNo] AS BankContract
FROM
  Schedule
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (Schedule.SchBnkClNO) Is Not Null
    )
  )
GROUP BY
  Schedule.SchID,
  [SchBnkClNO] & "-" & [SgrpBnkNo]
HAVING
  (
    (
      (Schedule.SchID)= 845
    )
  );

SELECT
  tbl_UserlogOUT.User_name,
  Max(tbl_UserlogOUT.Time_OUT_Stamp) AS MaxOfTime_OUT_Stamp
FROM
  tbl_UserlogOUT
GROUP BY
  tbl_UserlogOUT.User_name
HAVING
  (
    (
      (
        Max(tbl_UserlogOUT.Time_OUT_Stamp)
      )>= Date()
    )
  );

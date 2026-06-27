SELECT
  tbl_UserlogIN.User_name,
  Max(tbl_UserlogIN.Time_In_Stamp) AS MaxOfTime_In_Stamp
FROM
  tbl_UserlogIN
GROUP BY
  tbl_UserlogIN.User_name
HAVING
  (
    (
      (
        Max(tbl_UserlogIN.Time_In_Stamp)
      )>= Date()
    )
  );

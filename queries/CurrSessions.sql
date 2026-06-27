SELECT
  qry_login.User_name,
  qry_login.MaxOfTime_In_Stamp AS Time_In,
  IIf(
    [MaxOfTime_OUT_Stamp] < [MaxOfTime_In_Stamp],
    "", [MaxOfTime_OUT_Stamp]
  ) AS Time_Out,
  First(TBL_Employees.[Active User]) AS [FirstOfActive User]
FROM
  (
    qry_login
    LEFT JOIN qry_logout ON qry_login.User_name = qry_logout.User_name
  )
  LEFT JOIN TBL_Employees ON qry_logout.User_name = TBL_Employees.[User ID]
GROUP BY
  qry_login.User_name,
  qry_login.MaxOfTime_In_Stamp,
  IIf(
    [MaxOfTime_OUT_Stamp] < [MaxOfTime_In_Stamp],
    "", [MaxOfTime_OUT_Stamp]
  );

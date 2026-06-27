dbMemo "SQL" ="SELECT *\015\012FROM TFR_Work\015\012WHERE Assigned_Indicator='I' OR Assigned_In"
    "dicator <> '' AND FundDate >=#1/1/2024# AND FundDate <=#12/31/2024# Or FundDate "
    "Is Null;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End

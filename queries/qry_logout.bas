Operation =1
Option =0
Having ="(((Max(tbl_UserlogOUT.Time_OUT_Stamp))>=Date()))"
Begin InputTables
    Name ="tbl_UserlogOUT"
End
Begin OutputColumns
    Expression ="tbl_UserlogOUT.User_name"
    Alias ="MaxOfTime_OUT_Stamp"
    Expression ="Max(tbl_UserlogOUT.Time_OUT_Stamp)"
End
Begin Groups
    Expression ="tbl_UserlogOUT.User_name"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_UserlogOUT.User_name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MaxOfTime_OUT_Stamp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =509
    Top =227
    Right =1915
    Bottom =885
    Left =-1
    Top =-1
    Right =1374
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_UserlogOUT"
        Name =""
    End
End

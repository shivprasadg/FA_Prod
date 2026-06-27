Operation =1
Option =0
Having ="(((RemarketingReportAllUnits.LState) Is Not Null Or (RemarketingReportAllUnits.L"
    "State)<>''))"
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.LState"
End
Begin OrderBy
    Expression ="RemarketingReportAllUnits.LState"
    Flag =0
End
Begin Groups
    Expression ="RemarketingReportAllUnits.LState"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =580
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End

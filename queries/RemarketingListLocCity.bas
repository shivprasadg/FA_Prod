Operation =1
Option =0
Having ="(((RemarketingReportAllUnits.LCity) Is Not Null Or (RemarketingReportAllUnits.LC"
    "ity)<>\"\"))"
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.LCity"
End
Begin OrderBy
    Expression ="RemarketingReportAllUnits.LCity"
    Flag =0
End
Begin Groups
    Expression ="RemarketingReportAllUnits.LCity"
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
    Bottom =633
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =414
        Top =254
        Right =643
        Bottom =481
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End

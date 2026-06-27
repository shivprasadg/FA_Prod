Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns_Invoices_Drawn"
    Name ="DrawDowns"
End
Begin OutputColumns
    Expression ="DrawDowns_Invoices_Drawn.Invoice"
    Alias ="NumOfDraws"
    Expression ="Count(DrawDowns_Invoices_Drawn.DrawDownId)"
End
Begin Joins
    LeftTable ="DrawDowns_Invoices_Drawn"
    RightTable ="DrawDowns"
    Expression ="DrawDowns_Invoices_Drawn.DrawDownId = DrawDowns.DrawDownId"
    Flag =1
End
Begin OrderBy
    Expression ="DrawDowns_Invoices_Drawn.Invoice"
    Flag =0
End
Begin Groups
    Expression ="DrawDowns_Invoices_Drawn.Invoice"
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
        dbText "Name" ="DrawDowns_Invoices_Drawn.Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumOfDraws"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =86
    Top =113
    Right =1703
    Bottom =951
    Left =-1
    Top =-1
    Right =1593
    Bottom =624
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =77
        Top =8
        Right =446
        Bottom =306
        Top =0
        Name ="DrawDowns_Invoices_Drawn"
        Name =""
    End
    Begin
        Left =503
        Top =113
        Right =692
        Bottom =389
        Top =0
        Name ="DrawDowns"
        Name =""
    End
End

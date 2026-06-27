Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns"
End
Begin OutputColumns
    Expression ="DrawDowns.FAID"
End
Begin Groups
    Expression ="DrawDowns.FAID"
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
        dbText "Name" ="DrawDowns.FAID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =46
    Top =131
    Right =1557
    Bottom =854
    Left =-1
    Top =-1
    Right =1479
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =160
        Top =73
        Right =380
        Bottom =323
        Top =0
        Name ="DrawDowns"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((RmktBills.ReqDt)=Date()))"
Begin InputTables
    Name ="RmktBills"
End
Begin OutputColumns
    Expression ="RmktBills.[Invoice No]"
    Expression ="RmktBills.ReqDt"
    Expression ="RmktBills.BuyerRef"
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
        dbText "Name" ="RmktBills.ReqDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.[Invoice No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1595
    Bottom =855
    Left =-1
    Top =-1
    Right =1487
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =186
        Top =103
        Right =330
        Bottom =350
        Top =0
        Name ="RmktBills"
        Name =""
    End
End

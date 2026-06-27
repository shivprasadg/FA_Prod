Operation =1
Option =0
Begin InputTables
    Name ="ExitOptionTypes"
End
Begin OutputColumns
    Expression ="ExitOptionTypes.LeaseOptionType"
    Expression ="ExitOptionTypes.LeaseOptionId"
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
        dbText "Name" ="ExitOptionTypes.LeaseOptionType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionTypes.LeaseOptionId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1541
    Bottom =1021
    Left =-1
    Top =-1
    Right =1509
    Bottom =681
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ExitOptionTypes"
        Name =""
    End
End

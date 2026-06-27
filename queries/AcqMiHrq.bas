Operation =1
Option =0
Where ="(((Mileage.AquiMI)=Yes)) Or (((Mileage.AqiHr)=Yes))"
Begin InputTables
    Name ="Mileage"
End
Begin OutputColumns
    Expression ="Mileage.UnitID"
    Alias ="Expr1"
    Expression ="Mileage.MHDate"
    Alias ="Expr2"
    Expression ="Mileage.AquiMI"
    Alias ="Expr3"
    Expression ="Mileage.AqiHr"
    Alias ="Expr4"
    Expression ="Mileage.Mileage"
    Expression ="Mileage.Hours"
End
Begin OrderBy
    Expression ="Mileage.UnitID"
    Flag =0
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
        dbText "Name" ="Mileage.MHDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.AquiMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.AqiHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Hours"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Mileage"
        Name =""
    End
End

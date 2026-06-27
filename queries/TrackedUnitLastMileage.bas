Operation =1
Option =0
Where ="(((Mileage.UnitID) Is Not Null))"
Begin InputTables
    Name ="Mileage"
End
Begin OutputColumns
    Expression ="Mileage.UnitID"
    Alias ="MHDate"
    Expression ="Max(Mileage.MHDate)"
    Alias ="Odometer"
    Expression ="Max(Mileage.AquiMI)"
End
Begin OrderBy
    Expression ="Mileage.UnitID"
    Flag =0
End
Begin Groups
    Expression ="Mileage.UnitID"
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
        dbText "Name" ="Odometer"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MHDate"
        dbInteger "ColumnWidth" ="4020"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =195
    Top =141
    Right =1453
    Bottom =787
    Left =-1
    Top =-1
    Right =1226
    Bottom =375
    Left =15
    Top =0
    ColumnsShown =543
    Begin
        Left =176
        Top =68
        Right =320
        Bottom =212
        Top =0
        Name ="Mileage"
        Name =""
    End
End

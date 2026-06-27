Operation =1
Option =0
Where ="((([costco repairsNew].[Reefer Hours])>0))"
Begin InputTables
    Name ="costco repairsNew"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="[costco repairsNew].[Reefer Hours]"
    Alias ="Expr2"
    Expression ="[costco repairsNew].[Close DAte]"
    Alias ="UnitNo"
    Expression ="[costco repairsNew].[Unit #]"
    Alias ="Expr3"
    Expression ="[costco repairsNew].Unitref"
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
        dbText "Name" ="[costco repairsNew].[Reefer Hours]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[costco repairsNew].[Close DAte]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[costco repairsNew].Unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =183
    Top =141
    Right =1441
    Bottom =787
    Left =-1
    Top =-1
    Right =1226
    Bottom =422
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =262
        Top =-14
        Right =406
        Bottom =332
        Top =0
        Name ="costco repairsNew"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="Units_MakeModelsInUse_A"
End
Begin OutputColumns
    Alias ="UnitMake"
    Expression ="Units_MakeModelsInUse_A.Make"
    Alias ="UnitModel"
    Expression ="Units_MakeModelsInUse_A.Model"
    Alias ="Leased"
    Expression ="Max(Units_MakeModelsInUse_A.Leased)"
    Alias ="Tracked"
    Expression ="Max(Units_MakeModelsInUse_A.Tracked)"
End
Begin Groups
    Expression ="Units_MakeModelsInUse_A.Make"
    GroupLevel =0
    Expression ="Units_MakeModelsInUse_A.Model"
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
        dbText "Name" ="Leased"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_MakeModelsInUse_A.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_MakeModelsInUse_A.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitMake"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =50
    Top =104
    Right =1436
    Bottom =1002
    Left =-1
    Top =-1
    Right =1354
    Bottom =580
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =408
        Top =100
        Right =634
        Bottom =351
        Top =0
        Name ="Units_MakeModelsInUse_A"
        Name =""
    End
End

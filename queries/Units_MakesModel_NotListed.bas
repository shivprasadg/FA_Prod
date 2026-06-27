Operation =1
Option =0
Where ="(((LookUp_UnitModels.UnitModel) Is Null))"
Begin InputTables
    Name ="Units_MakeModelsInUse"
    Name ="LookUp_UnitModels"
End
Begin OutputColumns
    Expression ="Units_MakeModelsInUse.UnitMake"
    Expression ="Units_MakeModelsInUse.UnitModel"
    Expression ="Units_MakeModelsInUse.Leased"
    Expression ="Units_MakeModelsInUse.Tracked"
End
Begin Joins
    LeftTable ="Units_MakeModelsInUse"
    RightTable ="LookUp_UnitModels"
    Expression ="Units_MakeModelsInUse.UnitModel = LookUp_UnitModels.UnitModel"
    Flag =2
End
Begin OrderBy
    Expression ="Units_MakeModelsInUse.UnitMake"
    Flag =0
    Expression ="Units_MakeModelsInUse.UnitModel"
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
        dbText "Name" ="Units_MakeModelsInUse.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_MakeModelsInUse.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_MakeModelsInUse.Leased"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units_MakeModelsInUse.Tracked"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1682
    Bottom =991
    Left =-1
    Top =-1
    Right =1650
    Bottom =667
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =59
        Top =30
        Right =203
        Bottom =174
        Top =0
        Name ="Units_MakeModelsInUse"
        Name =""
    End
    Begin
        Left =331
        Top =126
        Right =475
        Bottom =270
        Top =0
        Name ="LookUp_UnitModels"
        Name =""
    End
End

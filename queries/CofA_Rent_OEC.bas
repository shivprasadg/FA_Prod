Operation =1
Option =0
Begin InputTables
    Name ="UnitOEC2"
    Name ="Bulk_Update_WorkTable"
End
Begin OutputColumns
    Expression ="UnitOEC2.AssetId"
    Expression ="UnitOEC2.TotalCost"
    Expression ="UnitOEC2.PerDiemRent"
    Expression ="UnitOEC2.SYNDRent"
    Expression ="UnitOEC2.LORent"
End
Begin Joins
    LeftTable ="Bulk_Update_WorkTable"
    RightTable ="UnitOEC2"
    Expression ="Bulk_Update_WorkTable.UnitID = UnitOEC2.AssetId"
    Flag =1
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
        dbText "Name" ="UnitOEC2.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.SYNDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.LORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1221
    Bottom =902
    Left =-1
    Top =-1
    Right =1189
    Bottom =545
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =234
        Top =148
        Right =378
        Bottom =535
        Top =0
        Name ="UnitOEC2"
        Name =""
    End
    Begin
        Left =40
        Top =53
        Right =184
        Bottom =197
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
End

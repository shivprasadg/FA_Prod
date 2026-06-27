Operation =1
Option =0
Begin InputTables
    Name ="UnitLevelTaxes"
    Name ="Bulk_Update_WorkTable"
End
Begin OutputColumns
    Expression ="UnitLevelTaxes.UnitID"
    Expression ="UnitLevelTaxes.UnitTaxes"
End
Begin Joins
    LeftTable ="UnitLevelTaxes"
    RightTable ="Bulk_Update_WorkTable"
    Expression ="UnitLevelTaxes.UnitID = Bulk_Update_WorkTable.UnitID"
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
        dbText "Name" ="UnitLevelTaxes.UnitTaxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLevelTaxes.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =650
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =215
        Top =123
        Right =359
        Bottom =283
        Top =0
        Name ="UnitLevelTaxes"
        Name =""
    End
    Begin
        Left =437
        Top =175
        Right =581
        Bottom =319
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
End

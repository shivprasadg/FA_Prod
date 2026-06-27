Operation =1
Option =0
Having ="(((Max(vw_UMP_Mileage.LogActivityDate))=#7/1/2015#) AND ((vw_UMP_Mileage.ClientG"
    "roupId)=60))"
Begin InputTables
    Name ="vw_UMP_Mileage"
End
Begin OutputColumns
    Alias ="LogActivityDate"
    Expression ="Max(vw_UMP_Mileage.LogActivityDate)"
    Expression ="vw_UMP_Mileage.ClientGroupId"
    Expression ="vw_UMP_Mileage.ClientId"
    Expression ="vw_UMP_Mileage.UnitId"
    Alias ="Expr1"
    Expression ="vw_UMP_Mileage.UnitNumber"
    Alias ="OBCDataStatus"
    Expression ="Min(vw_UMP_Mileage.OBCDataStatus)"
End
Begin Groups
    Expression ="vw_UMP_Mileage.ClientGroupId"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.ClientId"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.UnitId"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.UnitNumber"
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
        dbText "Name" ="LogActivityDate"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.UnitId"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1345
    Bottom =969
    Left =-1
    Top =-1
    Right =1313
    Bottom =544
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =81
        Top =48
        Right =396
        Bottom =469
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
End

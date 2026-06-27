Operation =1
Option =0
Where ="(((vw_UMP_Mileage.OBCDataStatus) Not Like \"*1792*\"))"
Having ="(((vw_UMP_Mileage.ClientGroupID)<>435))"
Begin InputTables
    Name ="vw_UMP_Mileage"
End
Begin OutputColumns
    Expression ="vw_UMP_Mileage.UnitID"
    Alias ="LogActivityDate"
    Expression ="Min(vw_UMP_Mileage.LogActivityDate)"
    Expression ="vw_UMP_Mileage.ClientID"
    Expression ="vw_UMP_Mileage.ClientGroupID"
End
Begin Groups
    Expression ="vw_UMP_Mileage.UnitID"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.ClientID"
    GroupLevel =0
    Expression ="vw_UMP_Mileage.ClientGroupID"
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
        dbText "Name" ="vw_UMP_Mileage.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1192
    Bottom =1021
    Left =-1
    Top =-1
    Right =1116
    Bottom =542
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =42
        Top =17
        Right =780
        Bottom =526
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="UnitLastOBCReport"
    Name ="vw_UMP_Mileage"
End
Begin OutputColumns
    Expression ="UnitLastOBCReport.UnitID"
    Alias ="LastRead"
    Expression ="vw_UMP_Mileage.LogActivityDate"
    Expression ="vw_UMP_Mileage.OdometerBegin"
    Expression ="vw_UMP_Mileage.OdometerEnd"
    Expression ="vw_UMP_Mileage.ClientID"
End
Begin Joins
    LeftTable ="UnitLastOBCReport"
    RightTable ="vw_UMP_Mileage"
    Expression ="UnitLastOBCReport.UnitID = vw_UMP_Mileage.UnitID"
    Flag =1
    LeftTable ="UnitLastOBCReport"
    RightTable ="vw_UMP_Mileage"
    Expression ="UnitLastOBCReport.LogActivityDate = vw_UMP_Mileage.LogActivityDate"
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
        dbText "Name" ="UnitLastOBCReport.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OdometerBegin"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OdometerEnd"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastRead"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1192
    Bottom =1021
    Left =-1
    Top =-1
    Right =1160
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =59
        Right =181
        Bottom =203
        Top =0
        Name ="UnitLastOBCReport"
        Name =""
    End
    Begin
        Left =253
        Top =88
        Right =522
        Bottom =407
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
End

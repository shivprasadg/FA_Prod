Operation =1
Option =0
Where ="(((vw_UMP_Mileage.ReportYear)=2016) AND ((vw_UMP_Mileage.ReportMonth)=8) AND ((v"
    "w_UMP_Mileage.ClientGroupID)=60) AND ((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.S"
    "tatus)=\"T\") AND ((vw_SixKeys.SubType)=\"DayCab\" Or (vw_SixKeys.SubType)=\"Sle"
    "eper\"))"
Begin InputTables
    Name ="vw_UMP_Mileage"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_UMP_Mileage.UnitID"
    Expression ="vw_UMP_Mileage.UnitNumber"
    Expression ="vw_UMP_Mileage.OBCUnitNumber"
    Expression ="vw_UMP_Mileage.LogActivityDate"
    Expression ="vw_UMP_Mileage.ReportYear"
    Expression ="vw_UMP_Mileage.ReportMonth"
    Expression ="vw_UMP_Mileage.ClientGroupID"
    Expression ="vw_UMP_Mileage.ClientID"
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.SubType"
End
Begin Joins
    LeftTable ="vw_UMP_Mileage"
    RightTable ="vw_SixKeys"
    Expression ="vw_UMP_Mileage.UnitID = vw_SixKeys.AssetID"
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
        dbText "Name" ="vw_UMP_Mileage.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ReportYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ReportMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SubType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1262
    Bottom =974
    Left =-1
    Top =-1
    Right =1230
    Bottom =565
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =170
        Top =102
        Right =528
        Bottom =496
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
    Begin
        Left =626
        Top =76
        Right =790
        Bottom =372
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End

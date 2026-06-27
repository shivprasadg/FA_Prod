Operation =1
Option =0
Where ="(((Units.UnitEstDelDate)<>[OpenVendorDataTemp].[UnitEstDelDate]))"
Begin InputTables
    Name ="OpenVendorDataTemp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Alias ="EstimatedDeliveryDate"
    Expression ="OpenVendorDataTemp.UnitEstDelDate"
    Alias ="PrevEstimatedDate"
    Expression ="Units.UnitEstDelDate"
End
Begin Joins
    LeftTable ="OpenVendorDataTemp"
    RightTable ="Units"
    Expression ="OpenVendorDataTemp.AssetID = Units.UnitID"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevEstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevEstimatedDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =127
    Top =98
    Right =1533
    Bottom =1096
    Left =-1
    Top =-1
    Right =1374
    Bottom =612
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =99
        Top =38
        Right =273
        Bottom =375
        Top =0
        Name ="OpenVendorDataTemp"
        Name =""
    End
    Begin
        Left =314
        Top =89
        Right =630
        Bottom =616
        Top =0
        Name ="Units"
        Name =""
    End
End

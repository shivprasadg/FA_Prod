Operation =1
Option =0
Where ="(((Units.UnitActualDelDate)<>[OpenVendorDataTemp].[UnitActualDelDate]))"
Begin InputTables
    Name ="OpenVendorDataTemp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Alias ="DeliveryDate"
    Expression ="OpenVendorDataTemp.UnitActualDelDate"
    Alias ="PrevDeliveryDate"
    Expression ="Units.UnitActualDelDate"
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
        dbText "Name" ="OpenVendorDataTemp.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenVendorDataTemp.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PONum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubmissionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevEstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevDelveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =156
    Top =43
    Right =1562
    Bottom =1041
    Left =-1
    Top =-1
    Right =1374
    Bottom =595
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

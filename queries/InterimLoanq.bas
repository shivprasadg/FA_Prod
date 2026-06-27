Operation =1
Option =0
Where ="(((Units.UnitIFAmt)>0.001) AND ((Clients.ClientGroupID)<>18) AND ((ChildPart_Cos"
    "t_OEC_summed.ExclFrGrpOEC)=False))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildPart_Cost_OEC_summed"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="UnitGroup"
    Expression ="SchGrp.SGrpNO"
    Alias ="Assignee"
    Expression ="SchGrp.Assignee_Bank"
    Alias ="LoanBank"
    Expression ="Units.UnitIFBank"
    Alias ="LoanDate"
    Expression ="Units.UnitIFDate"
    Alias ="LoanAmount"
    Expression ="Units.UnitIFAmt"
    Alias ="LoanType"
    Expression ="\"Interim\""
    Alias ="InterimExpiry"
    Expression ="SchGrp.SGrpExpDt"
    Alias ="PayoffDate"
    Expression ="Units.UnitPayOff"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="VIN"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Alias ="Cost"
    Expression ="[unitfinalcost]+Nz([PartsCost],0)"
    Alias ="Vendor_wParts"
    Expression ="[VendorCost]+Nz([PartsCost],0)"
    Expression ="Units.UnitIFDate"
    Expression ="Units.UnitIFBank"
    Expression ="Units.UnitIFAmt"
    Expression ="Units.UnitPayOff"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildPart_Cost_OEC_summed"
    Expression ="Units.UnitID = ChildPart_Cost_OEC_summed.UnitId"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPayOff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitGroup"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PayoffDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimExpiry"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LoanAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LoanType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LoanDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor_Parts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor_wParts"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =91
    Top =14
    Right =1629
    Bottom =660
    Left =-1
    Top =-1
    Right =1514
    Bottom =299
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =283
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =794
        Top =5
        Right =978
        Bottom =316
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1012
        Top =13
        Right =1189
        Bottom =200
        Top =0
        Name ="ChildPart_Cost_OEC_summed"
        Name =""
    End
End

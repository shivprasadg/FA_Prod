Operation =1
Option =0
Where ="(((NewULookupQ.ClientGroupID)<> 18))"
Begin InputTables
    Name ="Schedule"
    Name ="NewULookupQ"
    Name ="Rmkt"
    Name ="Clients"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="NewULookupQ.UnitID"
    Expression ="NewULookupQ.CompanyNm"
    Expression ="NewULookupQ.ClientShNm"
    Expression ="NewULookupQ.VendorSNm"
    Expression ="NewULookupQ.SalespersonName"
    Expression ="NewULookupQ.CompanyNm"
    Expression ="NewULookupQ.SPInitials"
    Expression ="NewULookupQ.ClientShNm"
    Expression ="NewULookupQ.VendorREF"
    Expression ="NewULookupQ.UnitVIN"
    Expression ="NewULookupQ.UnitPlateNum"
    Expression ="NewULookupQ.UnitPlateNumH"
    Expression ="NewULookupQ.UnitUnitNum"
    Expression ="NewULookupQ.UnitYr"
    Expression ="NewULookupQ.UnitMake"
    Expression ="NewULookupQ.UnitModel"
    Expression ="NewULookupQ.UnitType"
    Expression ="NewULookupQ.UnitVendorPytDate"
    Expression ="NewULookupQ.clientState"
    Expression ="NewULookupQ.clientCity"
    Alias ="Edit"
    Expression ="\"Edit\""
    Expression ="NewULookupQ.SGrpID"
    Expression ="NewULookupQ.GroupR"
    Expression ="NewULookupQ.UnitStatus"
    Expression ="NewULookupQ.LCity"
    Expression ="NewULookupQ.LState"
    Alias ="RBuyer"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchNo"
    Expression ="NewULookupQ.UnitInServ"
    Expression ="NewULookupQ.unitinservind"
End
Begin Joins
    LeftTable ="NewULookupQ"
    RightTable ="Rmkt"
    Expression ="NewULookupQ.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =2
    LeftTable ="NewULookupQ"
    RightTable ="SchGrp"
    Expression ="NewULookupQ.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewULookupQ.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.SalespersonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.CompanyNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitPlateNumH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.GroupR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RBuyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewULookupQ.unitinservind"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-96
    Top =40
    Right =694
    Bottom =1140
    Left =-1
    Top =-1
    Right =758
    Bottom =1019
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="NewULookupQ"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="SchGrp"
        Name =""
    End
End

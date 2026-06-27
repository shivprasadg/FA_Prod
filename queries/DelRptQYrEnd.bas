Operation =1
Option =0
Where ="(((Clients.ClientShNm) Not Like \"demo*\") AND ((Units.unitstatus)=\"b\") AND (("
    "MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Vendors"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Alias ="Key"
    Expression ="[schno] & [sgrpno] & [unitbuilddate] & [unitestdeldate]"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchFAID"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitID"
    Expression ="Units.unitvin"
    Expression ="Units.UnitID"
    Alias ="VIn4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unittype"
    Expression ="Units.unitstatus"
    Expression ="Units.UnitEstDelDate"
    Expression ="Units.UnitBuildDate"
    Expression ="Units.UnitUnitNum"
    Expression ="Schedule.SchShortDesc"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitMSORecd"
    Expression ="Units.UnitInvoiceNum"
    Expression ="Units.UnitInvoiceDate"
    Expression ="Units.UnitVendorPytDate"
    Expression ="SchGrp.BLCD"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorContact"
    Expression ="Vendors.VendorCity"
    Expression ="Vendors.VendorState"
    Expression ="Vendors.VendorZip"
    Expression ="Vendors.VendorPhone"
    Expression ="Vendors.VendorEmail"
    Expression ="Units.UnitInServ"
    Expression ="Units.UNITINSERVIND"
    Expression ="Units.RevDel"
    Expression ="Units.UnitDel"
    Expression ="Units.delstatus"
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
    LeftTable ="Vendors"
    RightTable ="Units"
    Expression ="Vendors.VendorID = Units.VendorREF"
    Flag =3
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
    Expression ="Units.unitvin"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Key"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIn4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorContact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITINSERVIND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RevDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.delstatus"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =122
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =15
        Top =4
        Right =159
        Bottom =148
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =195
        Top =9
        Right =339
        Bottom =153
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =392
        Top =10
        Right =536
        Bottom =154
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =570
        Top =7
        Right =714
        Bottom =151
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =961
        Top =12
        Right =1105
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =769
        Top =-4
        Right =913
        Bottom =140
        Top =0
        Name ="Units"
        Name =""
    End
End

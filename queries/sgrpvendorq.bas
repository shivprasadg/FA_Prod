Operation =1
Option =2
Having ="(((Clients.ClientShNm) Not Like \"demo*\") AND ((MstrLease.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Vendors"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchFAID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.SGrpFAID"
    Expression ="SchGrp.SGrpID"
    Expression ="Vendors.VendorID"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VENDORDBA"
    Expression ="Vendors.VendorSNm"
    Expression ="Vendors.VendorContact"
    Expression ="Vendors.VendorAddress"
    Expression ="Vendors.VendorCity"
    Expression ="Vendors.VendorState"
    Expression ="Vendors.VendorZip"
    Expression ="Vendors.VendorPhone"
    Expression ="Vendors.VendorFax"
    Expression ="Vendors.VendorPytInstructions"
    Expression ="Vendors.VendorEmail"
    Expression ="Vendors.VendorPmtType"
    Expression ="Vendors.VENDORPAYEE"
    Expression ="Vendors.VENDORABA"
    Expression ="Vendors.VENDORACH"
    Expression ="Vendors.VENDORBKACCT"
    Expression ="Vendors.VENDORBANK"
    Expression ="Vendors.VENDORBADD1"
    Expression ="Vendors.VEDNORBADD2"
    Expression ="Vendors.VENDORBCITY"
    Expression ="Vendors.VENDORBST"
    Expression ="Vendors.VENDORBZIP"
    Expression ="Vendors.VENDORPMTVER"
    Expression ="Vendors.VendorObsolete"
    Expression ="Vendors.VensorODt"
    Expression ="Vendors.VendorTaxID"
    Expression ="Units.UnitID"
    Expression ="Units.UnitID"
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
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin Groups
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchFAID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.SGrpFAID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Vendors.VendorID"
    GroupLevel =0
    Expression ="Vendors.VendorName"
    GroupLevel =0
    Expression ="Vendors.VENDORDBA"
    GroupLevel =0
    Expression ="Vendors.VendorSNm"
    GroupLevel =0
    Expression ="Vendors.VendorContact"
    GroupLevel =0
    Expression ="Vendors.VendorAddress"
    GroupLevel =0
    Expression ="Vendors.VendorCity"
    GroupLevel =0
    Expression ="Vendors.VendorState"
    GroupLevel =0
    Expression ="Vendors.VendorZip"
    GroupLevel =0
    Expression ="Vendors.VendorPhone"
    GroupLevel =0
    Expression ="Vendors.VendorFax"
    GroupLevel =0
    Expression ="Vendors.VendorPytInstructions"
    GroupLevel =0
    Expression ="Vendors.VendorEmail"
    GroupLevel =0
    Expression ="Vendors.VendorPmtType"
    GroupLevel =0
    Expression ="Vendors.VENDORPAYEE"
    GroupLevel =0
    Expression ="Vendors.VENDORABA"
    GroupLevel =0
    Expression ="Vendors.VENDORACH"
    GroupLevel =0
    Expression ="Vendors.VENDORBKACCT"
    GroupLevel =0
    Expression ="Vendors.VENDORBANK"
    GroupLevel =0
    Expression ="Vendors.VENDORBADD1"
    GroupLevel =0
    Expression ="Vendors.VEDNORBADD2"
    GroupLevel =0
    Expression ="Vendors.VENDORBCITY"
    GroupLevel =0
    Expression ="Vendors.VENDORBST"
    GroupLevel =0
    Expression ="Vendors.VENDORBZIP"
    GroupLevel =0
    Expression ="Vendors.VENDORPMTVER"
    GroupLevel =0
    Expression ="Vendors.VendorObsolete"
    GroupLevel =0
    Expression ="Vendors.VensorODt"
    GroupLevel =0
    Expression ="Vendors.VendorTaxID"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
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
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorContact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
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
        dbText "Name" ="Vendors.VendorFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPytInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPmtType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORPAYEE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBKACCT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBANK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBADD1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VEDNORBADD2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBZIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORPMTVER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VensorODt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
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
    Bottom =427
    Left =0
    Top =0
    ColumnsShown =543
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
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End

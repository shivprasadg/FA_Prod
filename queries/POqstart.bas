Operation =1
Option =0
Where ="(((IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]),[unitfaid"
    "]))=\"13-547\") AND ((MstrLease.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildParts"
    Name ="Vendors"
    Name ="Vendors"
    Alias ="Vendors_1"
    Name ="Locations"
End
Begin OutputColumns
    Alias ="KEY"
    Expression ="[vendorref] & [childvendorref] & [locationid] & [unitfinalcost] & [childfinalcos"
        "t]"
    Expression ="Units.UnitID"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorID"
    Expression ="Vendors.VendorSNm"
    Expression ="Vendors.VendorAddress"
    Expression ="Vendors.VendorCity"
    Expression ="Vendors.VendorState"
    Expression ="Vendors.VendorZip"
    Expression ="Vendors_1.vendorid"
    Expression ="Vendors_1.VendorName"
    Expression ="Vendors_1.VendorState"
    Expression ="Vendors_1.VendorZip"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Expression ="Units.unittype"
    Expression ="Units.UnitEstDelDate"
    Expression ="Units.unitfinalcost"
    Expression ="ChildParts.childyear"
    Expression ="ChildParts.childmake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.childtype"
    Expression ="ChildParts.childfinalcost"
    Alias ="FAID"
    Expression ="IIf(IsNull([unitfaid]),IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]),[unitfaid])"
    Expression ="Schedule.SchFAID"
    Expression ="SchGrp.SGrpFAID"
    Expression ="Units.UnitFAID"
    Expression ="MstrLease.MLOrig"
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
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
    LeftTable ="Units"
    RightTable ="Vendors"
    Expression ="Units.VendorREF = Vendors.VendorID"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="Vendors_1"
    Expression ="ChildParts.ChildVendorREF = Vendors_1.VendorID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitDeliveryLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
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
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitfinalcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KEY"
        dbInteger "ColumnWidth" ="3600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.schfaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitfaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorCity"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Vendors.vendorid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_1.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_1.VendorState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_1.VendorZip"
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
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childyear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childfinalcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_1.vendorid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
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
    Bottom =390
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
        Left =199
        Top =13
        Right =343
        Bottom =157
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =370
        Top =13
        Right =514
        Bottom =157
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =513
        Top =9
        Right =657
        Bottom =153
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =681
        Top =9
        Right =797
        Bottom =155
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =828
        Top =11
        Right =972
        Bottom =155
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =688
        Top =210
        Right =832
        Bottom =354
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =874
        Top =191
        Right =1018
        Bottom =335
        Top =0
        Name ="Vendors_1"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End

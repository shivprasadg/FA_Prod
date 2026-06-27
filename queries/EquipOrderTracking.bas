Operation =1
Option =0
Where ="(((MstrLease.MLOrig)=\"fa\") AND ((SchGrp.BLCD) Is Null)) OR (((MstrLease.MLOrig"
    ")=\"fa\") AND ((SchGrp.BLCD) Is Null)) OR (((MstrLease.MLOrig)=\"fa\") AND ((Sch"
    "Grp.BLCD) Is Null))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Vendors"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorAddress"
    Alias ="VendorCst"
    Expression ="[VendorCity] & \" \" & [vendorstate] & \" \" & [vendorzip]"
    Expression ="Vendors.VendorPhone"
    Alias ="Asset No"
    Expression ="Units.unitid"
    Alias ="PO FAID"
    Expression ="\"14-635\""
    Alias ="Pilot?"
    Expression ="IIf([UnitPilot]=True,\"Yes\",Null)"
    Alias ="VIN"
    Expression ="Units.Unitvin"
    Alias ="Deliver to"
    Expression ="IIf(IsNull([locsublessee]),[clientcompanyname],[locsublessee]) & \" \" & [locati"
        "onstreet] & \" \" & [LAddrLine1] & \" \" & [LCity] & \" \" & [LState] & \" \" & "
        "[LZip]"
    Alias ="Unit"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & \" \" & [unittype] & \" \" & IIf([axle]=\"TB"
        "D\",Null,[axle])"
    Alias ="Final Cost"
    Expression ="Units.UnitFinalCost"
    Alias ="Build DT"
    Expression ="Units.UnitBuildDate"
    Alias ="Est Delivery"
    Expression ="Units.unitestDelDate"
    Alias ="Del Date"
    Expression ="Units.UnitActualDelDate"
    Alias ="Engine Serial#"
    Expression ="Units.UnitEngineSerial"
    Alias ="Transmission Serial #"
    Expression ="Units.UnitTransSerial"
    Alias ="Drive Axle Serial #"
    Expression ="Units.UnitFASerial"
    Alias ="F Rear Axle Serial #"
    Expression ="Units.UnitRAFSerial"
    Alias ="R Rear Axle Serial #"
    Expression ="Units.UnitRARSerial"
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
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Vendors"
    Expression ="Units.VendorREF = Vendors.VendorID"
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
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
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="4215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPhone"
        dbLong "AggregateType" ="-1"
        dbText "InputMask" ="\\(999\") \"999\\-9999;;"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Build DT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="mm/dd/yy"
    End
    Begin
        dbText "Name" ="Del Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
        dbText "Format" ="mm/dd/yy"
    End
    Begin
        dbText "Name" ="Pilot?"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Asset No"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Unit"
        dbInteger "ColumnWidth" ="4455"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorCst"
        dbInteger "ColumnWidth" ="2790"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Est Delivery"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deliver to"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="7080"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Final Cost"
        dbInteger "ColumnWidth" ="1665"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Engine Serial#"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Transmission Serial #"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="F Rear Axle Serial #"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drive Axle Serial #"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R Rear Axle Serial #"
        dbInteger "ColumnWidth" ="1710"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PO FAID"
        dbInteger "ColumnWidth" ="1605"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =83
    Right =1348
    Bottom =732
    Left =-1
    Top =-1
    Right =1238
    Bottom =232
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =254
        Bottom =321
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =271
        Top =67
        Right =415
        Bottom =211
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =611
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
        Left =797
        Top =13
        Right =941
        Bottom =157
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1255
        Bottom =297
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =731
        Top =201
        Right =983
        Bottom =345
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =432
        Top =201
        Right =576
        Bottom =345
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End

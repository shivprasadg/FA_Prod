Operation =1
Option =0
Having ="(((MstrLease.MLOrig)=\"fa\") AND ((SchGrp.BLCD) Is Null))"
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
    Alias ="CountOfUnitID"
    Expression ="Count(Units.UnitID)"
    Alias ="Unit"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType] & IIf([axle]=\"TBD\",Null,\" \" & [axle])"
    Alias ="CompanyName"
    Expression ="Clients.clientCompanyName"
    Alias ="MinDeliveryDate"
    Expression ="Min(Units.UnitActualDelDate)"
    Alias ="MaxDeliveryDate"
    Expression ="Max(Units.UnitActualDelDate)"
    Alias ="FAID"
    Expression ="Schedule.schfaid"
    Expression ="Locations.LState"
    Expression ="Locations.LCity"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LZip"
    Expression ="Units.UnitFinalCost"
    Expression ="Locations.LocationID"
    Alias ="MinOfUnitEstDelDate"
    Expression ="Min(Units.UnitEstDelDate)"
    Alias ="MaxOfUnitestdeldate"
    Expression ="Max(Units.Unitestdeldate)"
    Expression ="Locations.LocSubLessee"
    Alias ="rptsec"
    Expression ="1"
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
Begin OrderBy
    Expression ="Locations.LState"
    Flag =0
    Expression ="Locations.LCity"
    Flag =0
    Expression ="Locations.LocationStreet"
    Flag =0
    Expression ="Units.UnitFinalCost"
    Flag =0
End
Begin Groups
    Expression ="Vendors.VendorName"
    GroupLevel =0
    Expression ="Vendors.VendorAddress"
    GroupLevel =0
    Expression ="[VendorCity] & \" \" & [vendorstate] & \" \" & [vendorzip]"
    GroupLevel =0
    Expression ="Vendors.VendorPhone"
    GroupLevel =0
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType] & IIf([axle]=\"TBD\",Null,\" \" & [axle])"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Schedule.schfaid"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LAddrLine1"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="Units.UnitFinalCost"
    GroupLevel =0
    Expression ="Locations.LocationID"
    GroupLevel =0
    Expression ="Locations.LocSubLessee"
    GroupLevel =0
    Expression ="1"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.sgrpfaid"
    GroupLevel =0
    Expression ="Units.unitfaid"
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
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="4215"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorPhone"
        dbInteger "ColumnWidth" ="1710"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbText "InputMask" ="\\(999\") \"999\\-9999;;"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
        dbInteger "ColumnWidth" ="1875"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbInteger "ColumnWidth" ="855"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbInteger "ColumnWidth" ="1665"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit"
        dbInteger "ColumnWidth" ="4080"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorCst"
        dbInteger "ColumnWidth" ="3570"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="4770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfUnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfUnitestdeldate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSubLessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rptsec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxDeliveryDate"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinDeliveryDate"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompanyName"
        dbInteger "ColumnWidth" ="2370"
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
    Bottom =267
    Left =0
    Top =0
    ColumnsShown =543
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
        Left =270
        Top =106
        Right =414
        Bottom =250
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

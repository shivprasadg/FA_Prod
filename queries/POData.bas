Operation =1
Option =0
Having ="(((SchGrp.PONum1)=Forms!GroupForm!PONum1) And ((MstrLease.MLOrig)=\"fa\"))"
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
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType] & IIf([axle]=\"TBD\",Null,\" \" & [axle])"
    Expression ="Clients.clientCompanyName"
    Alias ="Min Del Date"
    Expression ="Min(Units.UnitActualDelDate)"
    Alias ="Max Del Date"
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
    Expression ="SchGrp.PONum1"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.RebateUnitProjected"
    Expression ="SchGrp.POQuoteNums"
    Alias ="AttentionTo"
    Expression ="[Attention To]"
    Alias ="ChildCost"
    Expression ="DLookUp(\"ExtendedCost\",\"PO_ChildParts\",\"PoNumber='\" & [PONum1] & \"'\")"
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
    Flag =2
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
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
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
    Expression ="SchGrp.PONum1"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.RebateUnitProjected"
    GroupLevel =0
    Expression ="SchGrp.POQuoteNums"
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
        dbInteger "ColumnOrder" ="5"
        dbText "InputMask" ="\\(999\") \"999\\-9999;;"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnOrder" ="0"
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
        dbInteger "ColumnWidth" ="4455"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Min Del Date"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Max Del Date"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POQuoteNums"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AttentionTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalPOCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =94
    Top =38
    Right =1352
    Bottom =684
    Left =-1
    Top =-1
    Right =1240
    Bottom =182
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
        Left =609
        Top =14
        Right =753
        Bottom =158
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =797
        Top =13
        Right =948
        Bottom =274
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =991
        Top =157
        Right =1133
        Bottom =442
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =515
        Top =213
        Right =767
        Bottom =357
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =1039
        Top =49
        Right =1183
        Bottom =193
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End

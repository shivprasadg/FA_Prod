Operation =1
Option =0
Where ="(((Nz([unitinvoicenum],\"Need\"))=\"Need\") AND ((DateDiff(\"d\",Date(),[unitest"
    "deldate]))>-6) AND ((SchGrp.PONum1) Is Not Null) AND ((Clients.clienttype)<>\"de"
    "ad\") AND ((MstrLease.MLOrig)=\"FA\") AND ((CLASS.TITLED)=Yes))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="lookup_groupTypes"
    Name ="CLASS"
    Name ="Locations"
    Name ="Inv_MSOContactq2"
End
Begin OutputColumns
    Alias ="Regarding"
    Expression ="Clients.ClientShNm"
    Alias ="ClientName"
    Expression ="Clients.clientCompanyName"
    Alias ="FAID/(PO)"
    Expression ="Nz([sgrpfaid],[SchFAID])"
    Expression ="Units.unitvin"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="MYr"
    Expression ="Units.UnitYr"
    Alias ="Make"
    Expression ="Units.UnitMake"
    Alias ="Model"
    Expression ="Units.unitmodel"
    Alias ="Type"
    Expression ="Units.unittype"
    Alias ="SubType"
    Expression ="Units.unitSubType"
    Alias ="Est Delivelry Dt"
    Expression ="Units.UnitEstDelDate"
    Alias ="MSO Copy"
    Expression ="IIf(IsNull([unitmsocopy]) And IsNull([UnitMSORecd]),\"Needed\",Null)"
    Alias ="Invoice No"
    Expression ="Nz([unitinvoicenum],\"Need\")"
    Expression ="Locations.LCity"
    Alias ="Delivery City"
    Expression ="Locations.LCity"
    Alias ="Delivery St"
    Expression ="Locations.LState"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchNo"
    Expression ="Inv_MSOContactq2.VendorContact"
    Expression ="Inv_MSOContactq2.contactEmail"
    Expression ="Inv_MSOContactq2.VendorID"
    Alias ="GotMSO"
    Expression ="IIf(IsNull([UnitMSOCopy]),\"Need\",IIf(IsNull([UnitMSORecd]),\"Need\",\"Have\"))"
    Alias ="Days2Delivery"
    Expression ="DateDiff(\"d\",Date(),[unitestdeldate])"
    Expression ="Units.UnitMSORecd"
    Expression ="SchGrp.PONum1"
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
    Flag =2
    LeftTable ="Units"
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =2
    LeftTable ="lookup_groupTypes"
    RightTable ="CLASS"
    Expression ="lookup_groupTypes.grpClass = CLASS.CLASS"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Inv_MSOContactq2"
    Expression ="Units.VendorREF = Inv_MSOContactq2.VendorRef"
    Flag =2
End
Begin OrderBy
    Expression ="Nz([sgrpfaid],[SchFAID])"
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
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Est Delivelry Dt"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Invoice No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="MSO Copy"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="MYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Regarding"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="FAID/(PO)"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Delivery City"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Delivery St"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Inv_MSOContactq2.VendorContact"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Inv_MSOContactq2.contactEmail"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="ClientName"
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Inv_MSOContactq2.VendorID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Days2Delivery"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="GotMSO"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-316
    Top =51
    Right =942
    Bottom =710
    Left =-1
    Top =-1
    Right =1226
    Bottom =349
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =11
        Top =17
        Right =192
        Bottom =404
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =214
        Top =12
        Right =358
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =393
        Top =12
        Right =537
        Bottom =365
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =530
        Top =-11
        Right =674
        Bottom =133
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =758
        Top =15
        Right =942
        Bottom =350
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =998
        Top =6
        Right =1142
        Bottom =150
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
    Begin
        Left =1160
        Top =32
        Right =1304
        Bottom =176
        Top =0
        Name ="CLASS"
        Name =""
    End
    Begin
        Left =1001
        Top =157
        Right =1146
        Bottom =339
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =585
        Top =158
        Right =729
        Bottom =302
        Top =0
        Name ="Inv_MSOContactq2"
        Name =""
    End
End

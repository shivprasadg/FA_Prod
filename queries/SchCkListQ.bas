Operation =1
Option =0
Having ="(((MstrLease.MLNo)=[ENTER MASTER LEASE ID]) AND ((Schedule.SchNo)=[ENTER SCHED N"
    "O]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="lookup_States"
    Name ="Units"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLSent"
    Expression ="MstrLease.MLBack"
    Expression ="MstrLease.MLGuarantor"
    Expression ="MstrLease.MLGuarBack"
    Expression ="MstrLease.MLGuarSent"
    Expression ="MstrLease.MLInsExp"
    Expression ="MstrLease.MLINRecA"
    Expression ="MstrLease.MLINSBLANKET"
    Expression ="MstrLease.MLINSAccpt"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.UnitsInGroup"
    Expression ="SchGrp.OrigEquipCost"
    Expression ="Schedule.SchOut"
    Expression ="Schedule.SchIn"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitCOABack"
    Expression ="Units.UnitVINLocked"
    Expression ="Units.UnitInvoiceNum"
    Expression ="Units.UnitInvoiceDate"
    Expression ="Units.UnitFinalCost"
    Expression ="Units.UnitVendorPytDate"
    Expression ="Units.UnitMSORecd"
    Expression ="Units.UnitIRP"
    Expression ="Units.UnitTaxExempt"
    Expression ="Units.UnitTaxNotes"
    Expression ="Units.UnitTitleNumber"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
    Alias ="SumOfchildFinalCost"
    Expression ="Sum(ChildParts.childFinalCost)"
    Alias ="UDesc"
    Expression ="UCase([unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [unittype])"
    Expression ="Units.UnitTitleRecd"
    Expression ="Locations.LState"
    Expression ="Units.UnitID"
    Expression ="lookup_States.StTaxType"
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
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_1"
    Expression ="Units.UnitTitleLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_2"
    Expression ="Units.UnitRegLocationRef = Locations_2.LocationID"
    Flag =2
    LeftTable ="lookup_States"
    RightTable ="Locations"
    Expression ="lookup_States.StateAbbrev = Locations.LState"
    Flag =3
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
    Expression ="Units.UnitUnitNum"
    Flag =0
    Expression ="Units.UnitVIN"
    Flag =0
End
Begin Groups
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="MstrLease.MLSent"
    GroupLevel =0
    Expression ="MstrLease.MLBack"
    GroupLevel =0
    Expression ="MstrLease.MLGuarantor"
    GroupLevel =0
    Expression ="MstrLease.MLGuarBack"
    GroupLevel =0
    Expression ="MstrLease.MLGuarSent"
    GroupLevel =0
    Expression ="MstrLease.MLInsExp"
    GroupLevel =0
    Expression ="MstrLease.MLINRecA"
    GroupLevel =0
    Expression ="MstrLease.MLINSBLANKET"
    GroupLevel =0
    Expression ="MstrLease.MLINSAccpt"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="SchGrp.OrigEquipCost"
    GroupLevel =0
    Expression ="Schedule.SchOut"
    GroupLevel =0
    Expression ="Schedule.SchIn"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.UnitUnitNum"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="Units.UnitCOABack"
    GroupLevel =0
    Expression ="Units.UnitVINLocked"
    GroupLevel =0
    Expression ="Units.UnitInvoiceNum"
    GroupLevel =0
    Expression ="Units.UnitInvoiceDate"
    GroupLevel =0
    Expression ="Units.UnitFinalCost"
    GroupLevel =0
    Expression ="Units.UnitVendorPytDate"
    GroupLevel =0
    Expression ="Units.UnitMSORecd"
    GroupLevel =0
    Expression ="Units.UnitIRP"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Units.UnitTaxNotes"
    GroupLevel =0
    Expression ="Units.UnitTitleNumber"
    GroupLevel =0
    Expression ="Units.UnitMake"
    GroupLevel =0
    Expression ="Units.UnitModel"
    GroupLevel =0
    Expression ="Units.UnitType"
    GroupLevel =0
    Expression ="UCase([unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & [unittype])"
    GroupLevel =0
    Expression ="Units.UnitTitleRecd"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="lookup_States.StTaxType"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Right([unitvin],6)"
    GroupLevel =0
    Expression ="Units.UnitUnitNum"
    GroupLevel =0
    Expression ="Units.UnitVIN"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Units.UnitTaxNotes"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRP"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLINSAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLINRecA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLINSBLANKET"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarantor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLInsExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfchildFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =938
    Bottom =676
    Left =-1
    Top =-1
    Right =885
    Bottom =315
    Left =307
    Top =0
    ColumnsShown =543
    Begin
        Left =-269
        Top =5
        Right =-77
        Bottom =149
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =-66
        Top =11
        Right =78
        Bottom =155
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =95
        Top =15
        Right =239
        Bottom =159
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =264
        Top =14
        Right =408
        Bottom =158
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =815
        Top =43
        Right =959
        Bottom =187
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =426
        Top =6
        Right =598
        Bottom =146
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =629
        Top =8
        Right =773
        Bottom =152
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =419
        Top =171
        Right =563
        Bottom =315
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =699
        Top =167
        Right =843
        Bottom =311
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =216
        Top =176
        Right =360
        Bottom =320
        Top =0
        Name ="ChildParts"
        Name =""
    End
End

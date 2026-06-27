Operation =1
Option =0
Where ="(((SchGrp.LeaseType)<>\"loan\") AND ((Clients.ClientShNm) Not Like \"demo*\") AN"
    "D ((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"c\" Or (Units.unitstatus)=\""
    "b\"))"
Having ="(((Units.UnitType) Is Not Null And (Units.UnitType)<>\"misc\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Alias ="yr"
    Expression ="IIf(IsNull([unitacceptdate]),DatePart(\"yyyy\",Date()),DatePart(\"yyyy\",[UnitAc"
        "ceptdate]))"
    Alias ="UType"
    Expression ="Units.Unitsubtype"
    Alias ="CountOfUnitID"
    Expression ="Count(Units.UnitID)"
    Alias ="grpClass"
    Expression ="Units.UnitType"
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
    RightTable ="lookup_groupTypes"
    Expression ="Units.UnitType = lookup_groupTypes.groupType"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="IIf(IsNull([unitacceptdate]),DatePart(\"yyyy\",Date()),DatePart(\"yyyy\",[UnitAc"
        "ceptdate]))"
    Flag =0
    Expression ="Units.Unitsubtype"
    Flag =0
    Expression ="Count(Units.UnitID)"
    Flag =0
End
Begin Groups
    Expression ="IIf(IsNull([unitacceptdate]),DatePart(\"yyyy\",Date()),DatePart(\"yyyy\",[UnitAc"
        "ceptdate]))"
    GroupLevel =0
    Expression ="Units.Unitsubtype"
    GroupLevel =0
    Expression ="Units.UnitType"
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
        dbText "Name" ="yr"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UType"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="grpClass"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =115
    Top =43
    Right =1373
    Bottom =689
    Left =-1
    Top =-1
    Right =1240
    Bottom =202
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
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End

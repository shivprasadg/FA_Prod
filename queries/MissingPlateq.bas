Operation =1
Option =0
Where ="(((IIf([SchTitleToTrust]=True,\"Yes\",\"No\"))=\"yes\") AND ((IIf([SchTitltoFA]="
    "True,\"Yes\",\"No\"))=\"no\") AND ((Units.UnitAcceptDate) Is Not Null) AND ((Uni"
    "ts.UnitDocAccpt)=Yes) AND ((Units.UnitInServ) Is Not Null And (Units.UnitInServ)"
    ">#12/31/2010#) AND ((Units.unitinservind)=\"a\") AND ((Clients.ClientGroupID)<> "
    "18) AND ((Units.UnitPlateNum) Is Null) AND ((MstrLease.MLOrig)=\"fa\") AND ((Uni"
    "ts.unitstatus)=\"a\" Or (Units.unitstatus)=\"tr\" Or (Units.unitstatus)=\"c\")) "
    "OR (((IIf([SchTitleToTrust]=True,\"Yes\",\"No\"))=\"no\") AND ((IIf([SchTitltoFA"
    "]=True,\"Yes\",\"No\"))=\"yes\") AND ((Units.UnitAcceptDate) Is Not Null) AND (("
    "Units.UnitDocAccpt)=Yes) AND ((Units.UnitInServ) Is Not Null And (Units.UnitInSe"
    "rv)>#12/31/2010#) AND ((Units.unitinservind)=\"a\") AND ((Clients.ClientGroupID)"
    "<> 18) AND ((Units.UnitPlateNum) Is Null) AND ((MstrLease.MLOrig)=\"fa\") AND (("
    "Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"tr\" Or (Units.unitstatus)=\"c\""
    ")) OR (((IIf([SchTitleToTrust]=True,\"Yes\",\"No\"))=\"yes\") AND ((IIf([SchTitl"
    "toFA]=True,\"Yes\",\"No\"))=\"no\") AND ((Units.UnitAcceptDate) Is Not Null And "
    "(Units.UnitAcceptDate)>#12/31/2010#) AND ((Units.UnitDocAccpt)=No) AND ((Clients"
    ".ClientGroupID)<> 18) AND ((Units.UnitPlateNum) Is Null) AND ((MstrLease.MLOrig)"
    "=\"fa\") AND ((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"tr\" Or (Units.un"
    "itstatus)=\"c\")) OR (((IIf([SchTitleToTrust]=True,\"Yes\",\"No\"))=\"no\") AND "
    "((IIf([SchTitltoFA]=True,\"Yes\",\"No\"))=\"yes\") AND ((Units.UnitAcceptDate) I"
    "s Not Null And (Units.UnitAcceptDate)>#12/31/2010#) AND ((Units.UnitDocAccpt)=No"
    ") AND ((Clients.ClientGroupID)<> 18) AND ((Units.UnitPlateNum) Is Null) AND ((Ms"
    "trLease.MLOrig)=\"fa\") AND ((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"tr"
    "\" Or (Units.unitstatus)=\"c\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Alias ="TITLETrust?"
    Expression ="IIf([SchTitleToTrust]=True,\"Yes\",\"No\")"
    Alias ="TiteltoFA?"
    Expression ="IIf([SchTitltoFA]=True,\"Yes\",\"No\")"
    Expression ="SchGrp.SGrpNO"
    Expression ="Units.unitunitnum"
    Expression ="Units.unitvin"
    Expression ="Units.Unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unittype"
    Expression ="Locations.LState"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitDocAccpt"
    Expression ="Units.UnitInServ"
    Expression ="Units.unitinservind"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.ClientGroupID"
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
    Expression ="Units.UnitTitleLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Units.Unityr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinservind"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TITLETrust?"
        dbInteger "ColumnWidth" ="1845"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TiteltoFA?"
        dbInteger "ColumnWidth" ="1470"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
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
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End

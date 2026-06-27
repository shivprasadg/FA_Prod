Operation =1
Option =0
Where ="(((Clients.ClientGroupID)<> 18) AND ((Units.unittype)<>\"Yard Tractor\") AND ((U"
    "nits.unittasent) Is Not Null) AND ((Units.UnitTitleNumber) Is Null) AND ((MstrLe"
    "ase.MLOrig)=\"FA\")) OR (((Clients.ClientGroupID)<> 18) AND ((Units.unittype)<>\""
    "Yard Tractor\") AND ((Units.UnitTitleNumber) Is Null) AND ((UCase([UnitStatus]))"
    "=\"A\") AND ((MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Alias ="Customer"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.ClientGroupID"
    Alias ="Lease NO"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Alias ="Assignee"
    Expression ="SchGrp.Assignee_Bank"
    Alias ="Titled to Trust?"
    Expression ="IIf([SchTitleToTrust]=True,\"Yes\",\"No\")"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="VIN"
    Expression ="UCase([unitvin])"
    Alias ="Year"
    Expression ="Units.unityr"
    Alias ="Make"
    Expression ="Units.unitmake"
    Alias ="Model"
    Expression ="Units.UnitModel"
    Alias ="Type"
    Expression ="Units.unittype"
    Alias ="SubType"
    Expression ="Units.UnitSubType"
    Alias ="App Sent"
    Expression ="Units.unittasent"
    Alias ="Title Received"
    Expression ="Units.UnitTitleRecd"
    Alias ="Title No"
    Expression ="Units.UnitTitleNumber"
    Alias ="Status"
    Expression ="UCase([UnitStatus])"
    Alias ="Title ST"
    Expression ="Locations.LState"
    Alias ="Days o/s"
    Expression ="DateDiff(\"d\",[unittasent],Date())"
    Alias ="As Of"
    Expression ="Date()"
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
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
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
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease NO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="App Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Received"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Title ST"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Days o/s"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Titled to Trust?"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="As Of"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_groupTypes.grpClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.grpClass"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =61
    Top =17
    Right =1418
    Bottom =770
    Left =-1
    Top =-1
    Right =1325
    Bottom =392
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
        Name ="Locations"
        Name =""
    End
End

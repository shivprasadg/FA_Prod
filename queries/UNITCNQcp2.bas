Operation =1
Option =0
Where ="(((Units.unitstatus)=\"a\" Or (Units.unitstatus)=\"c\" Or (Units.unitstatus)=\"t"
    "r\") AND ((MstrLease.MLOrig)=\"fa\"))"
Having ="(((Clients.clientID)=forms!MgmtMenu!SelCp) And ((Clients.ClientShNm) Not Like \""
    "demo*\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Alias ="yr"
    Expression ="DatePart(\"yyyy\",[UnitAcceptDate])"
    Alias ="UType"
    Expression ="IIf(Left([UnitType],6)=\"Daycab\",\"Daycab\",[unittype])"
    Alias ="CountOfunitid"
    Expression ="Count(Units.unitid)"
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
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
End
Begin OrderBy
    Expression ="DatePart(\"yyyy\",[UnitAcceptDate])"
    Flag =0
    Expression ="IIf(Left([UnitType],6)=\"Daycab\",\"Daycab\",[unittype])"
    Flag =0
    Expression ="Count(Units.unitid)"
    Flag =0
End
Begin Groups
    Expression ="DatePart(\"yyyy\",[UnitAcceptDate])"
    GroupLevel =0
    Expression ="IIf(Left([UnitType],6)=\"Daycab\",\"Daycab\",[unittype])"
    GroupLevel =0
    Expression ="Clients.ClientGroupID"
    GroupLevel =0
    Expression ="Clients.ClientGroupID"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="IIf(IsNull(ClientGroupID),[clientcompanyname],ClientGroupID)"
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
        dbText "Name" ="UType"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[Client Group]"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfunitid"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =46
        Top =13
        Right =190
        Bottom =157
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =211
        Top =11
        Right =355
        Bottom =155
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =444
        Top =12
        Right =588
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =636
        Top =12
        Right =780
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =865
        Top =6
        Right =1009
        Bottom =150
        Top =0
        Name ="Units"
        Name =""
    End
End

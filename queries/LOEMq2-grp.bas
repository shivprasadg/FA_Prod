Operation =1
Option =0
Where ="(((IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt]))>DateAdd(\"d\",Date(),-60"
    ")) AND ((Notes.CallType)=\"e-mail\") AND ((Notes.noteUser)=[loinitial]))"
Begin InputTables
    Name ="Notes"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="CP"
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.ClientShNm"
    Alias ="CountOfnoteid"
    Expression ="Count(Notes.noteid)"
    Alias ="Last15E"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<=Date() And IIf(IsNu"
        "ll([NoteEctDt]),[notedatetime],[noteectdt])>DateAdd(\"d\",Date(),-15),1,0))"
    Alias ="Last30E"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<DateAdd(\"d\",Date()"
        ",-15) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>=DateAdd(\"d\",Dat"
        "e(),-30),1,0))"
    Alias ="Last45E"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<=DateAdd(\"d\",Date("
        "),-30) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>DateAdd(\"d\",Dat"
        "e(),-45),1,0))"
    Alias ="Last60E"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<=DateAdd(\"d\",Date("
        "),-45) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>DateAdd(\"d\",Dat"
        "e(),-60),1,0))"
    Expression ="Clients.loinitial"
End
Begin Joins
    LeftTable ="Notes"
    RightTable ="Clients"
    Expression ="Notes.clientRef = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientGroupID"
    Flag =0
End
Begin Groups
    Expression ="Clients.ClientGroupID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Clients.loinitial"
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
        dbText "Name" ="CountOfnoteid"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last15E"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last45E"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last60E"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CP"
        dbInteger "ColumnWidth" ="3405"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last30E"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.loinitial"
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
    Bottom =86
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =425
        Top =-8
        Right =569
        Bottom =136
        Top =0
        Name ="Notes"
        Name =""
    End
    Begin
        Left =266
        Top =2
        Right =410
        Bottom =146
        Top =0
        Name ="Clients"
        Name =""
    End
End

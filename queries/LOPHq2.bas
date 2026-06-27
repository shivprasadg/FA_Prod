Operation =1
Option =0
Where ="(((IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt]))>=DateAdd(\"d\",Date(),-6"
    "0)) AND ((Notes.CallType)=\"phone call\"))"
Begin InputTables
    Name ="Notes"
    Name ="Clients"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="Notes.noteUser"
    Expression ="Clients.clientID"
    Expression ="Clients.clientcompanyname"
    Alias ="CountOfnoteid"
    Expression ="Count(Notes.noteid)"
    Alias ="Last15p"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<Date() And IIf(IsNul"
        "l([NoteEctDt]),[notedatetime],[noteectdt])>=DateAdd(\"d\",Date(),-15),1,0))"
    Alias ="Last30p"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<DateAdd(\"d\",Date()"
        ",-15) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>=DateAdd(\"d\",Dat"
        "e(),-30),1,0))"
    Alias ="Last45p"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<DateAdd(\"d\",Date()"
        ",-30) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>=DateAdd(\"d\",Dat"
        "e(),-45),1,0))"
    Alias ="Last60p"
    Expression ="Sum(IIf(IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])<DateAdd(\"d\",Date()"
        ",-45) And IIf(IsNull([NoteEctDt]),[notedatetime],[noteectdt])>=DateAdd(\"d\",Dat"
        "e(),-60),1,0))"
End
Begin Joins
    LeftTable ="Notes"
    RightTable ="Clients"
    Expression ="Notes.clientRef = Clients.clientID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.LOinitial = Salesperson.SPInitials"
    Flag =1
End
Begin OrderBy
    Expression ="Notes.noteUser"
    Flag =0
    Expression ="Clients.clientcompanyname"
    Flag =0
End
Begin Groups
    Expression ="Notes.noteUser"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Clients.clientcompanyname"
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
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Clients.clientcompanyname"
        dbInteger "ColumnWidth" ="2445"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last15p"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last30p"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last45p"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last60p"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.noteUser"
        dbInteger "ColumnOrder" ="2"
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
    Bottom =115
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
    Begin
        Left =56
        Top =-13
        Right =200
        Bottom =131
        Top =0
        Name ="Salesperson"
        Name =""
    End
End

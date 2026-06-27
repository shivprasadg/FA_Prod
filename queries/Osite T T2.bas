Operation =1
Option =0
Having ="(((Clients.ClientType)=\"target\" Or (Clients.ClientType) Like \"2nd*\") AND ((N"
    "otes.CallType)=\"on-site\"))"
Begin InputTables
    Name ="Notes"
    Name ="Clients"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="Salesperson.SPInitials"
    Expression ="Clients.ClientType"
    Expression ="Clients.clientCompanyName"
    Alias ="MaxOfnoteDateTime"
    Expression ="Max(Notes.noteDateTime)"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.SalespersonREF = Salesperson.SalespersonID"
    Flag =1
    LeftTable ="Notes"
    RightTable ="Clients"
    Expression ="Notes.clientRef = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Salesperson.SPInitials"
    Flag =0
    Expression ="Clients.ClientType"
    Flag =1
End
Begin Groups
    Expression ="Salesperson.SPInitials"
    GroupLevel =0
    Expression ="Clients.ClientType"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Notes.CallType"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfnoteDateTime"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
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
    Bottom =363
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Notes"
        Name =""
    End
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
        Left =42
        Top =149
        Right =186
        Bottom =293
        Top =0
        Name ="Salesperson"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((Clients.clientType) Like \"*target*\") AND ((Contacts.contactRTitle)=\"4\" Or"
    " (Contacts.contactRTitle)=\"5\") AND ((Contacts.ContactObsolete)=No))"
Begin InputTables
    Name ="Clients"
    Name ="Contacts"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Expression ="Salesperson.SPInitials"
    Expression ="Clients.clientType"
    Alias ="Nm"
    Expression ="[contactfirst] & \" \" & [contactlast]"
    Expression ="Contacts.contactTitle"
    Expression ="Contacts.contactRTitle"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactPhone2"
    Expression ="Contacts.contactCell"
    Expression ="Contacts.contactEmail"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Contacts"
    Expression ="Clients.clientID = Contacts.clientRef"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.SalespersonREF = Salesperson.SalespersonID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Contacts.contactLast"
    Flag =0
    Expression ="Contacts.contactFirst"
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
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbInteger "ColumnWidth" ="3135"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Nm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactRTitle"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrimaryPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCell"
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
    Bottom =301
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
        Right =438
        Bottom =164
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =166
        Top =173
        Right =310
        Bottom =317
        Top =0
        Name ="Salesperson"
        Name =""
    End
End

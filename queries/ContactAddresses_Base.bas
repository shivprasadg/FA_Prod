Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.contactID"
    Alias ="ContactAddress"
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & Nz([ContactAddress1],[clientAddress1])"
        " & Chr(13) & Chr(10) & IIf(IsNull(Nz([ContactAddress2],[clientAddress2])),Null,N"
        "z([ContactAddress2],[clientAddress1]) & Chr(13) & Chr(10)) & Nz([ContactCity],[c"
        "lientCity]) & \" \" & Nz([Contacts].[LState],[Clientstate]) & \" \" & Left(Nz([c"
        "ontactZip],[clientzip]),5)"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactEmail"
    Alias ="contactPhone"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientID"
    Expression ="Clients.clientAddress1"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Contacts"
    Expression ="Clients.clientID = Contacts.clientRef"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="840"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactAddress"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="contactPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1336
    Bottom =969
    Left =-1
    Top =-1
    Right =1318
    Bottom =618
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =212
        Bottom =433
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =639
        Bottom =433
        Top =0
        Name ="Contacts"
        Name =""
    End
End

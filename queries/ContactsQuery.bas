Operation =1
Option =0
Where ="(((Clients.clientType)<>\"dead\") AND ((Contacts.ContactObsolete)=No))"
Begin InputTables
    Name ="Contacts"
    Name ="client_lookup"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="client_lookup.CompanyAndDBA"
    Expression ="Contacts.contactClassification"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactEmail"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactPhoneExt"
    Expression ="Contacts.contactCell"
    Expression ="Contacts.contactPhone2"
    Expression ="Contacts.contactPrefix"
    Expression ="Contacts.contactFirst"
    Expression ="Contacts.contactMiddle"
    Expression ="Contacts.contactLast"
    Expression ="Contacts.contactSuffix"
    Expression ="Contacts.contactTitle"
    Expression ="Contacts.contactOptOut"
    Expression ="Clients.clientType"
End
Begin Joins
    LeftTable ="Contacts"
    RightTable ="client_lookup"
    Expression ="Contacts.clientRef = client_lookup.clientID"
    Flag =1
    LeftTable ="client_lookup"
    RightTable ="Clients"
    Expression ="client_lookup.clientID = Clients.clientID"
    Flag =1
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
        dbText "Name" ="client_lookup.CompanyAndDBA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3210"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Contacts.contactPrefix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFirst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactMiddle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactLast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSuffix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrimaryPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhoneExt"
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
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOptOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients.clientType"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =547
        Top =18
        Right =772
        Bottom =248
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =369
        Top =14
        Right =513
        Bottom =158
        Top =0
        Name ="client_lookup"
        Name =""
    End
    Begin
        Left =160
        Top =11
        Right =304
        Bottom =155
        Top =0
        Name ="Clients"
        Name =""
    End
End

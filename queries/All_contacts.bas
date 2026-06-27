Operation =1
Option =0
Begin InputTables
    Name ="client_lookup"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="client_lookup.CompanyAndDBA"
    Expression ="Contacts.clientRef"
    Expression ="Contacts.contactID"
    Expression ="Contacts.contactCreateDate"
    Expression ="Contacts.contactEditDate"
    Expression ="Contacts.contactClassification"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactPrefix"
    Expression ="Contacts.contactFirst"
    Expression ="Contacts.contactMiddle"
    Expression ="Contacts.contactLast"
    Expression ="Contacts.contactSuffix"
    Expression ="Contacts.contactTitle"
    Expression ="Contacts.contactSalutation"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactPhoneExt"
    Expression ="Contacts.contactPhone2"
    Expression ="Contacts.contactCell"
    Expression ="Contacts.contactEmail"
    Expression ="Contacts.contactOptOut"
    Expression ="Contacts.contactNote"
    Expression ="Contacts.contactAddress1"
    Expression ="Contacts.contactAddress2"
    Expression ="Contacts.contactCity"
    Expression ="Contacts.stateRef"
    Expression ="Contacts.contactZip"
    Expression ="Contacts.contactCountry"
    Expression ="Contacts.LocationName"
    Expression ="Contacts.contactFax"
    Expression ="Contacts.ContactObsolete"
    Expression ="Contacts.ContactObsoletedt"
End
Begin Joins
    LeftTable ="client_lookup"
    RightTable ="Contacts"
    Expression ="client_lookup.clientID = Contacts.clientRef"
    Flag =1
End
Begin OrderBy
    Expression ="client_lookup.CompanyAndDBA"
    Flag =0
    Expression ="Contacts.contactClassification"
    Flag =0
    Expression ="Contacts.contactFirst"
    Flag =0
    Expression ="Contacts.contactLast"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
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
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactLast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFirst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCreateDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Contacts.contactEditDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrefix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactMiddle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSuffix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSalutation"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOptOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="client_lookup.CompanyAndDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsoletedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
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
    Right =1240
    Bottom =84
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="client_lookup"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
End

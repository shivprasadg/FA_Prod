Operation =1
Option =0
Begin InputTables
    Name ="ContactsLinkedToLocation"
    Name ="Contacts"
    Name ="Locations"
    Name ="Clients"
    Name ="ContactsLinkedToLocationTypes"
End
Begin OutputColumns
    Expression ="ContactsLinkedToLocation.ContactID"
    Expression ="ContactsLinkedToLocation.LocationID"
    Expression ="ContactsLinkedToLocationTypes.LinkType"
    Alias ="Contact"
    Expression ="[contactFirst] & \" \" & [contactLast]"
    Expression ="Contacts.contactEmail"
    Expression ="Locations.Locationsn"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="State"
    Expression ="Locations.LState"
    Expression ="Locations.LocationType"
    Expression ="ContactsLinkedToLocation.LinkTypeID"
End
Begin Joins
    LeftTable ="ContactsLinkedToLocation"
    RightTable ="Contacts"
    Expression ="ContactsLinkedToLocation.ContactID = Contacts.contactID"
    Flag =1
    LeftTable ="ContactsLinkedToLocation"
    RightTable ="Locations"
    Expression ="ContactsLinkedToLocation.LocationID = Locations.LocationID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Clients"
    Expression ="Locations.ClientREF = Clients.clientID"
    Flag =1
    LeftTable ="ContactsLinkedToLocation"
    RightTable ="ContactsLinkedToLocationTypes"
    Expression ="ContactsLinkedToLocation.LinkTypeID = ContactsLinkedToLocationTypes.LinkTypeID"
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
dbMemo "Filter" ="([ContactsLinkedToLocationQuery].[LocationID]=952)"
Begin
    Begin
        dbText "Name" ="Contact"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ContactsLinkedToLocation.contactId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactsLinkedToLocation.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactsLinkedToLocationTypes.LinkType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactsLinkedToLocation.LinkTypeID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =113
    Top =82
    Right =1427
    Bottom =1016
    Left =-1
    Top =-1
    Right =1282
    Bottom =514
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ContactsLinkedToLocation"
        Name =""
    End
    Begin
        Left =603
        Top =238
        Right =747
        Bottom =382
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =222
        Top =178
        Right =401
        Bottom =534
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =428
        Top =46
        Right =573
        Bottom =192
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =429
        Top =250
        Right =573
        Bottom =394
        Top =0
        Name ="ContactsLinkedToLocationTypes"
        Name =""
    End
End

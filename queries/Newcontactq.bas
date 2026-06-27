Operation =1
Option =0
Where ="(((Clients.clientCompanyName)<>\"FLEET ADVANTAGE\" And (Clients.clientCompanyNam"
    "e) Not Like \"demo*\") And ((Contacts.contactCreateDate)>=forms!call_report_para"
    "m!startdate And (Contacts.contactCreateDate)<forms!call_report_param!enddate) An"
    "d ((Clients.clienttype)<>\"dead\")) Or (((Clients.clientCompanyName)<>\"FLEET AD"
    "VANTAGE\" And (Clients.clientCompanyName) Not Like \"demo*\") And ((Contacts.Las"
    "tChangeDt)>=forms!call_report_param!startdate And (Contacts.LastChangeDt)<forms!"
    "call_report_param!enddate)) Or (((Contacts.ContactObsoletedt)>=forms!call_report"
    "_param!startdate And (Contacts.ContactObsoletedt)<forms!call_report_param!enddat"
    "e))"
Begin InputTables
    Name ="Clients"
    Name ="Contacts"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientAddress1"
    Expression ="Clients.clientAddress2"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientZip"
    Expression ="Clients.clientMainPhone"
    Expression ="Clients.clientFax"
    Expression ="Contacts.contactCreateDate"
    Expression ="Contacts.LastChangeDt"
    Alias ="NM"
    Expression ="[contactlast] & \" \" & [contactsuffix] & \", \" & [contactfirst] & \" \" & [con"
        "tactmiddle] & \" \" & [contacttitle]"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactPhoneExt"
    Expression ="Contacts.contactPhone2"
    Expression ="Contacts.contactFax"
    Expression ="Contacts.contactCell"
    Expression ="Contacts.contactEmail"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LAddrLine3"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Contacts.contactLast"
    Expression ="Contacts.contactFirst"
    Expression ="Contacts.contactMiddle"
    Expression ="Clients.CLIENTTYPE"
    Expression ="Clients.clienttype"
    Alias ="srtr"
    Expression ="IIf(IsNull([lastchangedt]),\"New\",\"Changed\")"
    Expression ="Contacts.ContactObsoletedt"
    Expression ="Contacts.ContactObsolete"
End
Begin Joins
    LeftTable ="Contacts"
    RightTable ="Locations"
    Expression ="Contacts.cLocationID = Locations.LocationID"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Contacts"
    Expression ="Clients.clientID = Contacts.clientRef"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientMainPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCreateDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NM"
        dbInteger "ColumnWidth" ="6600"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Contacts.contactFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCell"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbInteger "ColumnWidth" ="3945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
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
        dbText "Name" ="Contacts.contactMiddle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.CLIENTTYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LastChangeDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1027"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="srtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsoletedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LCity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1359
    Bottom =793
    Left =-1
    Top =-1
    Right =1335
    Bottom =229
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =230
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =485
        Bottom =255
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =533
        Top =12
        Right =677
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End

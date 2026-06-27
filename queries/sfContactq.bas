Operation =1
Option =0
Where ="(((Clients.clientType)=\"client\" Or (Clients.clientType)=\"guarantor\" Or (Clie"
    "nts.clientType) Like \"*target\") AND ((Contacts.ContactObsolete)=No))"
Begin InputTables
    Name ="Clients"
    Name ="Contacts"
    Name ="Locations"
End
Begin OutputColumns
    Alias ="Unique Id"
    Expression ="Contacts.contactID"
    Alias ="ACCOUNT Unique Id"
    Expression ="Clients.clientID"
    Alias ="LASTNAME"
    Expression ="Contacts.contactLast"
    Alias ="FIRSTNAME"
    Expression ="Contacts.contactFirst"
    Alias ="SALUTATION"
    Expression ="Contacts.contactPrefix"
    Alias ="OTHERSTREET"
    Expression ="Locations.LocationStreet"
    Alias ="OTHERCITY"
    Expression ="Locations.LCity"
    Alias ="OTHERSTATE"
    Expression ="Locations.LState"
    Alias ="OTHERPOSTALCODE"
    Expression ="Locations.LZip"
    Alias ="OTHERCOUNTRY"
    Expression ="Locations.LCountry"
    Alias ="MAILINGSTREET"
    Expression ="Null"
    Alias ="MAILINGCITY"
    Expression ="Null"
    Alias ="MAILINGSTATE"
    Expression ="Null"
    Alias ="MAILINGPOSTALCODE"
    Expression ="Null"
    Alias ="MAILINGCOUNTRY"
    Expression ="Null"
    Alias ="PHONE"
    Expression ="Contacts.contactPrimaryPhone"
    Alias ="FAX"
    Expression ="Contacts.contactFax"
    Alias ="MOBILEPHONE"
    Expression ="Contacts.contactCell"
    Alias ="HOMEPHONE"
    Expression ="Null"
    Alias ="OTHERPHONE"
    Expression ="Contacts.contactPhone2"
    Alias ="ASSISTANTPHONE"
    Expression ="Contacts.ContactAsstPH"
    Alias ="REPORTSTO"
    Expression ="0"
    Alias ="EMAIL"
    Expression ="Contacts.contactEmail"
    Alias ="TITLE"
    Expression ="Contacts.contactTitle"
    Alias ="DEPARTMENT"
    Expression ="Contacts.contactClassification"
    Alias ="ASSISTANTNAME"
    Expression ="[CONTACTASSTFNM] & \"  \" & [CONTACTASSTLNM]"
    Alias ="LEADSOURCE"
    Expression ="Null"
    Alias ="BIRTHDATE"
    Expression ="Null"
    Alias ="DESCRIPTON"
    Expression ="Null"
    Alias ="OWNER"
    Expression ="Null"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Contacts"
    Expression ="Clients.clientID = Contacts.clientRef"
    Flag =1
    LeftTable ="Contacts"
    RightTable ="Locations"
    Expression ="Contacts.cLocationID = Locations.LocationID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientID"
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
        dbText "Name" ="Unique Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ACCOUNT Unique Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LASTNAME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FIRSTNAME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SALUTATION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERSTREET"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERSTATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERPOSTALCODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERCOUNTRY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAILINGSTREET"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAILINGCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAILINGSTATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAILINGPOSTALCODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAILINGCOUNTRY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MOBILEPHONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OTHERPHONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HOMEPHONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ASSISTANTPHONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="REPORTSTO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EMAIL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TITLE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4755"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DEPARTMENT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ASSISTANTNAME"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LEADSOURCE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BIRTHDATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DESCRIPTON"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OWNER"
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
    Right =1234
    Bottom =186
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
        Right =423
        Bottom =188
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =471
        Top =12
        Right =674
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End

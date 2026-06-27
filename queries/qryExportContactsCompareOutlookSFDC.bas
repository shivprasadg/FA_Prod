Operation =1
Option =0
Begin InputTables
    Name ="Contacts"
    Name ="Clients"
    Name ="Locations"
    Name ="ClientGroups"
End
Begin OutputColumns
    Alias ="Pre"
    Expression ="Contacts.contactPrefix"
    Alias ="FirstN"
    Expression ="Trim([contactFirst] & IIf(Nz([contactMiddle],\"\")<>\" \",\" \" & [contactMiddle"
        "],\"\"))"
    Alias ="LastN"
    Expression ="Trim([contactLast] & IIf(Nz([contactSuffix],\"\")<>\" \",\" \" & [contactSuffix]"
        ",\"\"))"
    Alias ="Comp"
    Expression ="[clientCompanyName] & IIf(Nz([ClientGroupName],\"\")<>\"\",\" / \" & [ClientGrou"
        "pName],\"\")"
    Alias ="Cls"
    Expression ="Contacts.contactClassification"
    Alias ="Title"
    Expression ="Contacts.contactTitle"
    Alias ="Address"
    Expression ="[LAddrLine1] & IIf(Nz([LAddrLine2],\"\")<>\"\",\", \" & [LAddrLine2],\"\") & IIf"
        "(Nz([LAddrLine3],\"\")<>\"\",\", \" & [LAddrLine3],\"\")"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="St"
    Expression ="Locations.LState"
    Alias ="Zip"
    Expression ="Locations.LZip"
    Alias ="Country"
    Expression ="Locations.LCountry"
    Alias ="OthStr"
    Expression ="Null"
    Alias ="OthCity"
    Expression ="Null"
    Alias ="OthSt"
    Expression ="Null"
    Alias ="OthZip"
    Expression ="Null"
    Alias ="AsstPhone"
    Expression ="Contacts.ContactAsstPH"
    Alias ="Fax"
    Expression ="Contacts.contactFax"
    Alias ="Phone"
    Expression ="[contactPrimaryPhone] & IIf(Nz([contactPhoneExt],\"\")<>\" \",\" ext. \" & [cont"
        "actPhoneExt],\"\")"
    Alias ="Home"
    Expression ="Null"
    Alias ="Mobile"
    Expression ="Contacts.contactCell"
    Alias ="OthPhone"
    Expression ="Contacts.contactPhone2"
    Alias ="AsstName"
    Expression ="Trim(IIf(Nz([ContactAsstFnm],\"\")<>\"\",[ContactAsstFnm],\"\") & IIf(Nz([Contac"
        "tAsstLnm],\"\")<>\"\",\" \" & [ContactAsstLnm],\"\"))"
    Alias ="dob"
    Expression ="Null"
    Alias ="cat"
    Expression ="Null"
    Alias ="Email"
    Expression ="Contacts.contactemail"
    Alias ="mgr"
    Expression ="Null"
    Alias ="Note"
    Expression ="IIf(Nz([ContactObsolete],0)=1,\"OBSOLETE \",\"\") & IIf(Nz([DoNotCall],0)=1,\"DO"
        " NOT CALL \",\"\") & IIf(Nz([contactOptOut],0)=1,\"OPT OUT \",\"\") & [ContactNo"
        "te]"
    Expression ="Contacts.contactID"
End
Begin Joins
    LeftTable ="Contacts"
    RightTable ="Clients"
    Expression ="Contacts.clientRef = Clients.clientID"
    Flag =2
    LeftTable ="Contacts"
    RightTable ="Locations"
    Expression ="Contacts.cLocationID = Locations.LocationID"
    Flag =2
    LeftTable ="Clients"
    RightTable ="ClientGroups"
    Expression ="Clients.ClientGroupId = ClientGroups.ClientGroupId"
    Flag =2
End
Begin OrderBy
    Expression ="Nz([ClientGroupName],\"zzzz\" & [clientcompanyname])"
    Flag =0
    Expression ="Contacts.contactLast"
    Flag =0
    Expression ="Contacts.contactFirst"
    Flag =0
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Contacts.contactEmail"
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
        dbText "Name" ="Pre"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AsstPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OthPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="mgr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dob"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Email"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cls"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OthZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mobile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Comp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Home"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AsstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OthSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Address"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OthStr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OthCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Country"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1482
    Bottom =1038
    Left =-1
    Top =-1
    Right =1458
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="ClientGroups"
        Name =""
    End
End

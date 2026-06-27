Operation =1
Option =0
Where ="(((VendorContacts.ContactObsolete)=No))"
Begin InputTables
    Name ="Vendors"
    Name ="VendorContacts"
End
Begin OutputColumns
    Expression ="Vendors.VendorSNm"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorAddress"
    Expression ="Vendors.VendorCity"
    Expression ="Vendors.VendorState"
    Expression ="Vendors.VendorZip"
    Alias ="LastName"
    Expression ="VendorContacts.contactLast"
    Alias ="FristName"
    Expression ="VendorContacts.contactFirst"
    Alias ="Middle"
    Expression ="VendorContacts.contactMiddle"
    Alias ="Title"
    Expression ="VendorContacts.contactTitle"
    Alias ="Phone"
    Expression ="VendorContacts.contactPrimaryPhone"
    Alias ="Email"
    Expression ="VendorContacts.contactEmail"
    Alias ="Mobil"
    Expression ="VendorContacts.contactCell"
    Alias ="Address1"
    Expression ="VendorContacts.contactAddress1"
    Alias ="Addrees2"
    Expression ="VendorContacts.contactAddress2"
    Alias ="City"
    Expression ="VendorContacts.contactCity"
    Alias ="State"
    Expression ="VendorContacts.LState"
    Alias ="Zip"
    Expression ="VendorContacts.contactZip"
    Expression ="Vendors.VENDORBANK"
    Alias ="OPT_In_Out"
    Expression ="IIf([VCEMoptout]=False,\"IN\",\"OUT\")"
End
Begin Joins
    LeftTable ="Vendors"
    RightTable ="VendorContacts"
    Expression ="Vendors.VendorID = VendorContacts.VendorRef"
    Flag =2
End
Begin OrderBy
    Expression ="Vendors.VendorSNm"
    Flag =0
    Expression ="VendorContacts.contactLast"
    Flag =0
    Expression ="VendorContacts.contactFirst"
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
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VENDORBANK"
        dbInteger "ColumnWidth" ="3600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Addrees2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mobil"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Email"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Middle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FristName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OPT_In_Out"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-370
    Top =5
    Right =1302
    Bottom =798
    Left =-1
    Top =-1
    Right =1640
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =0
        Right =241
        Bottom =237
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =285
        Top =19
        Right =577
        Bottom =278
        Top =0
        Name ="VendorContacts"
        Name =""
    End
End

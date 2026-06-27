Operation =1
Option =0
Begin InputTables
    Name ="ClientContacts_Base"
    Name ="ATF_ContactsActive_Base"
End
Begin OutputColumns
    Expression ="ATF_ContactsActive_Base.ClientID"
    Expression ="ATF_ContactsActive_Base.MLID"
    Alias ="SignLevel"
    Expression ="IIf([SignerType]=\"Master Lease\",9,IIf([SignerType]=\"Schedule\",8,IIf([SignerT"
        "ype]=\"Schedule\",1,0)))"
    Expression ="ATF_ContactsActive_Base.SignerType"
    Expression ="ATF_ContactsActive_Base.SchID"
    Expression ="ATF_ContactsActive_Base.ContactID"
    Expression ="ClientContacts_Base.FullName"
    Expression ="ClientContacts_Base.Title"
    Expression ="ClientContacts_Base.Phone"
    Expression ="ClientContacts_Base.Mobile"
    Expression ="ClientContacts_Base.HasEmail"
    Alias ="Sorter"
    Expression ="IIf([SignerType]=\"Master Lease\",1,IIf([SignerType]=\"Guarantor\",2,IIf([Signer"
        "Type]=\"Schedule\",3,IIf([SignerType]=\"COA\",4,10))))"
End
Begin Joins
    LeftTable ="ATF_ContactsActive_Base"
    RightTable ="ClientContacts_Base"
    Expression ="ATF_ContactsActive_Base.ContactID = ClientContacts_Base.contactID"
    Flag =1
End
Begin Groups
    Expression ="ATF_ContactsActive_Base.ClientID"
    GroupLevel =0
    Expression ="ATF_ContactsActive_Base.MLID"
    GroupLevel =0
    Expression ="ATF_ContactsActive_Base.SignerType"
    GroupLevel =0
    Expression ="ATF_ContactsActive_Base.SchID"
    GroupLevel =0
    Expression ="ATF_ContactsActive_Base.ContactID"
    GroupLevel =0
    Expression ="ClientContacts_Base.FullName"
    GroupLevel =0
    Expression ="ClientContacts_Base.Title"
    GroupLevel =0
    Expression ="ClientContacts_Base.Phone"
    GroupLevel =0
    Expression ="ClientContacts_Base.Mobile"
    GroupLevel =0
    Expression ="ClientContacts_Base.HasEmail"
    GroupLevel =0
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
        dbText "Name" ="ClientContacts_Base.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive_Base.ContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientContacts_Base.FullName"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientContacts_Base.Title"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientContacts_Base.Mobile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive_Base.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientContacts_Base.HasEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive_Base.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive_Base.SignerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive_Base.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sorter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SignLevel"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =103
    Top =17
    Right =1536
    Bottom =950
    Left =-1
    Top =-1
    Right =1415
    Bottom =537
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =289
        Top =80
        Right =433
        Bottom =401
        Top =0
        Name ="ClientContacts_Base"
        Name =""
    End
    Begin
        Left =35
        Top =15
        Right =241
        Bottom =308
        Top =0
        Name ="ATF_ContactsActive_Base"
        Name =""
    End
End

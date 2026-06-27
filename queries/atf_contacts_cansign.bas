Operation =6
Option =0
Begin InputTables
    Name ="ATF_ContactsActive"
End
Begin OutputColumns
    Expression ="ATF_ContactsActive.ContactID"
    GroupLevel =2
    Expression ="ATF_ContactsActive.FullName"
    GroupLevel =2
    Expression ="ATF_ContactsActive.Title"
    GroupLevel =2
    Expression ="ATF_ContactsActive.SignerType"
    GroupLevel =1
    Alias ="SignerType"
    Expression ="First(ATF_ContactsActive.SignerType)"
End
Begin OrderBy
    Expression ="ATF_ContactsActive.SignerType"
    Flag =1
End
Begin Groups
    Expression ="ATF_ContactsActive.ContactID"
    GroupLevel =2
    Expression ="ATF_ContactsActive.FullName"
    GroupLevel =2
    Expression ="ATF_ContactsActive.Title"
    GroupLevel =2
    Expression ="ATF_ContactsActive.SignerType"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="ATF_ContactsActive.ContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Guarantor Type"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master Lease Type"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.SignerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COA Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule Type"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Guarantor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master Lease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =223
    Top =25
    Right =1486
    Bottom =839
    Left =-1
    Top =-1
    Right =1245
    Bottom =652
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =160
        Top =137
        Right =524
        Bottom =433
        Top =0
        Name ="ATF_ContactsActive"
        Name =""
    End
End

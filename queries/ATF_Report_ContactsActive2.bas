Operation =1
Option =0
Begin InputTables
    Name ="ATF_ContactsActive"
    Name ="ATF_Report_Active_ContactControl"
End
Begin OutputColumns
    Alias ="Sorter"
    Expression ="ATF_ContactsActive.Sorter"
    Expression ="ATF_ContactsActive.ClientID"
    Expression ="ATF_ContactsActive.MLID"
    Alias ="SignerType"
    Expression ="ATF_ContactsActive.SignerType"
    Expression ="ATF_ContactsActive.SchID"
    Expression ="ATF_ContactsActive.ContactID"
    Expression ="ATF_ContactsActive.FullName"
    Expression ="ATF_ContactsActive.Title"
    Expression ="ATF_ContactsActive.Phone"
    Expression ="ATF_ContactsActive.Mobile"
    Expression ="ATF_ContactsActive.HasEmail"
End
Begin Joins
    LeftTable ="ATF_Report_Active_ContactControl"
    RightTable ="ATF_ContactsActive"
    Expression ="ATF_Report_Active_ContactControl.Sorter = ATF_ContactsActive.Sorter"
    Flag =1
    LeftTable ="ATF_Report_Active_ContactControl"
    RightTable ="ATF_ContactsActive"
    Expression ="ATF_Report_Active_ContactControl.SchID = ATF_ContactsActive.SchID"
    Flag =1
    LeftTable ="ATF_Report_Active_ContactControl"
    RightTable ="ATF_ContactsActive"
    Expression ="ATF_Report_Active_ContactControl.ContactID = ATF_ContactsActive.ContactID"
    Flag =1
End
Begin OrderBy
    Expression ="ATF_ContactsActive.Sorter"
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
dbMemo "Filter" ="([ATF_Report_ContactsActive2].[SchID]=932)"
Begin
    Begin
        dbText "Name" ="ATF_ContactsActive.FullName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sorter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.ContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.HasEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ContactsActive.Mobile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SignerType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =94
    Top =195
    Right =1571
    Bottom =851
    Left =-1
    Top =-1
    Right =1459
    Bottom =218
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =221
        Top =20
        Right =365
        Bottom =164
        Top =0
        Name ="ATF_ContactsActive"
        Name =""
    End
    Begin
        Left =31
        Top =20
        Right =175
        Bottom =164
        Top =0
        Name ="ATF_Report_Active_ContactControl"
        Name =""
    End
End

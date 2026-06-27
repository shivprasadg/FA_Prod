Operation =1
Option =0
Begin InputTables
    Name ="ATF_Report_ContactsActive"
End
Begin OutputColumns
    Expression ="ATF_Report_ContactsActive.SchID"
    Alias ="Sorter"
    Expression ="Min(ATF_Report_ContactsActive.Sorter)"
    Alias ="ContactID"
    Expression ="ATF_Report_ContactsActive.ContactID"
End
Begin Groups
    Expression ="ATF_Report_ContactsActive.SchID"
    GroupLevel =0
    Expression ="ATF_Report_ContactsActive.ContactID"
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
        dbText "Name" ="ATF_Report_ContactsActive.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sorter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1533
    Bottom =1026
    Left =-1
    Top =-1
    Right =1515
    Bottom =709
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =116
        Top =56
        Right =260
        Bottom =200
        Top =0
        Name ="ATF_Report_ContactsActive"
        Name =""
    End
End

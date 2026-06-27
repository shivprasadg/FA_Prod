Operation =1
Option =0
Having ="(((Location_contacts.contactEmail) Is Not Null) AND ((Location_contacts.contactC"
    "lassification) Like \"*TR*\"))"
Begin InputTables
    Name ="Location_contacts"
End
Begin OutputColumns
    Alias ="BankId"
    Expression ="Location_contacts.clientRef"
    Alias ="TCName"
    Expression ="Location_contacts.contactFullName"
    Alias ="TCEmail"
    Expression ="Location_contacts.contactEmail"
    Expression ="Location_contacts.contactClassification"
End
Begin OrderBy
    Expression ="Location_contacts.clientRef"
    Flag =0
    Expression ="Location_contacts.contactEmail"
    Flag =0
End
Begin Groups
    Expression ="Location_contacts.clientRef"
    GroupLevel =0
    Expression ="Location_contacts.contactFullName"
    GroupLevel =0
    Expression ="Location_contacts.contactEmail"
    GroupLevel =0
    Expression ="Location_contacts.contactClassification"
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
dbMemo "Filter" ="([BankTrustEmails].[BankId]=4209)"
Begin
    Begin
        dbText "Name" ="BankId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location_contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TCEmail"
        dbInteger "ColumnWidth" ="3645"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TCName"
        dbInteger "ColumnWidth" ="3645"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =181
    Top =95
    Right =2270
    Bottom =974
    Left =-1
    Top =-1
    Right =2065
    Bottom =358
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =268
        Top =73
        Right =464
        Bottom =560
        Top =0
        Name ="Location_contacts"
        Name =""
    End
End

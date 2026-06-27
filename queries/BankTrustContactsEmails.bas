Operation =1
Option =0
Where ="(((LC.contactEmail) Is Not Null) AND ((LC.contactClassification) Like \"TR*\" Or"
    " (LC.contactClassification) Like \"*Trust*\"))"
Begin InputTables
    Name ="Location_contacts"
    Alias ="LC"
End
Begin OutputColumns
    Alias ="BankId"
    Expression ="LC.clientRef"
    Alias ="TrustContactName"
    Expression ="LC.contactFullName"
    Alias ="TrustContactEmail"
    Expression ="LC.contactEmail"
    Expression ="LC.contactClassification"
End
Begin OrderBy
    Expression ="LC.clientRef"
    Flag =0
    Expression ="LC.contactEmail"
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
        dbText "Name" ="BankId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrustContactName"
        dbInteger "ColumnWidth" ="2685"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrustContactEmail"
        dbInteger "ColumnWidth" ="3330"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LC.contactClassification"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =20
    Top =83
    Right =2109
    Bottom =962
    Left =-1
    Top =-1
    Right =2065
    Bottom =460
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="LC"
        Name =""
    End
End

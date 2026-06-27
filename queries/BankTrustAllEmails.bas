Operation =1
Option =0
Where ="(((IIf([contactClassification]=\"Primary Trust Incumbent\",[contactEmail],Null))"
    " Is Not Null)) OR (((IIf([contactClassification]=\"TR Incumbent\",[contactEmail]"
    ",Null)) Is Not Null)) OR (((IIf([contactClassification]=\"TR Incumbent-Contact\""
    ",[contactEmail],Null)) Is Not Null)) OR (((IIf([contactClassification]=\"TR Cont"
    "act\",[contactEmail],Null)) Is Not Null))"
Begin InputTables
    Name ="Location_contacts"
End
Begin OutputColumns
    Alias ="BankId"
    Expression ="Location_contacts.clientRef"
    Alias ="PrimaryTrustIncumbentEmail"
    Expression ="IIf([contactClassification]=\"Primary Trust Incumbent\",[contactEmail],Null)"
    Alias ="TrustIncumbentName"
    Expression ="IIf([contactClassification]=\"TR Incumbent\",[contactFullName],Null)"
    Alias ="TrustIncumbentEmail"
    Expression ="IIf([contactClassification]=\"TR Incumbent\",[contactEmail],Null)"
    Alias ="TRIncumbentContactName"
    Expression ="IIf([contactClassification]=\"TR Incumbent-Contact\",[contactFullName],Null)"
    Alias ="TRIncumbentContactEmail"
    Expression ="IIf([contactClassification]=\"TR Incumbent-Contact\",[contactEmail],Null)"
    Alias ="TrustContactName"
    Expression ="IIf([contactClassification]=\"TR Contact\",[contactFullName],Null)"
    Alias ="TrustContactEmail"
    Expression ="IIf([contactClassification]=\"TR Contact\",[contactEmail],Null)"
End
Begin OrderBy
    Expression ="IIf([contactClassification]=\"Primary Trust Incumbent\",[contactEmail],Null)"
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
        dbText "Name" ="BankId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrustContactName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2685"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TRIncumbentContactEmail"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TrustContactEmail"
        dbInteger "ColumnWidth" ="3330"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TRIncumbentContactName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2910"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TrustIncumbentName"
        dbInteger "ColumnWidth" ="3645"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrimaryTrustIncumbentEmail"
        dbInteger "ColumnWidth" ="3660"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrustIncumbentEmail"
        dbInteger "ColumnWidth" ="3645"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =191
    Top =-5
    Right =2280
    Bottom =874
    Left =-1
    Top =-1
    Right =2057
    Bottom =562
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =82
        Top =0
        Right =384
        Bottom =516
        Top =0
        Name ="Location_contacts"
        Name =""
    End
End

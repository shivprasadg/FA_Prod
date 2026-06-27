Operation =1
Option =0
Where ="(((VendorContacts.INV_MSO)=True))"
Begin InputTables
    Name ="VendorContacts"
End
Begin OutputColumns
    Expression ="VendorContacts.contactID"
    Expression ="VendorContacts.VendorRef"
    Expression ="VendorContacts.contactEmail"
    Alias ="VendorContact"
    Expression ="[contactfirst] & [contactmiddle] & \" \" & [contactlast]"
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
        dbText "Name" ="VendorContacts.contactID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorContacts.VendorRef"
        dbInteger "ColumnWidth" ="1425"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorContacts.contactEmail"
        dbInteger "ColumnWidth" ="3840"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorContact"
        dbInteger "ColumnWidth" ="4380"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
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
    Right =1226
    Bottom =390
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =491
        Top =48
        Right =759
        Bottom =391
        Top =0
        Name ="VendorContacts"
        Name =""
    End
End

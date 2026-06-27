Operation =1
Option =0
Where ="(((Contacts.contactInc)<>No) AND ((Contacts.ContactObsolete)<>Yes))"
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.clientRef"
    Alias ="ClientSignTitle"
    Expression ="Contacts.contactTitle"
    Alias ="ClientSignFullName"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactID"
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
        dbText "Name" ="Contacts.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientSignTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientSignFullName"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =181
    Top =46
    Right =1545
    Bottom =861
    Left =-1
    Top =-1
    Right =1332
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =560
        Top =24
        Right =954
        Bottom =481
        Top =0
        Name ="Contacts"
        Name =""
    End
End

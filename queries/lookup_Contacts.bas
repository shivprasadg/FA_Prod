Operation =1
Option =0
Where ="(((Contacts.clientRef)=[Forms]![Deals_view]![clientRef]))"
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.contactID"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactLast"
    Expression ="Contacts.contactFirst"
    Expression ="Contacts.contactTitle"
    Expression ="Contacts.clientRef"
End
Begin OrderBy
    Expression ="Contacts.contactLast"
    Flag =0
    Expression ="Contacts.contactFirst"
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
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactLast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFirst"
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
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =319
        Bottom =248
        Top =0
        Name ="Contacts"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.clientRef"
    Expression ="Contacts.contactClassification"
    Expression ="Contacts.contactPrefix"
    Expression ="Contacts.contactFirst"
    Expression ="Contacts.contactMiddle"
    Expression ="Contacts.contactLast"
    Expression ="Contacts.contactSuffix"
    Expression ="Contacts.contactTitle"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactPhoneExt"
    Expression ="Contacts.contactPhone2"
    Expression ="Contacts.contactCell"
    Expression ="Contacts.contactEmail"
    Expression ="Contacts.contactOptOut"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactNote"
End
Begin OrderBy
    Expression ="Contacts.contactClassification"
    Flag =0
    Expression ="Contacts.contactFirst"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
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
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSuffix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOptOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFirst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactMiddle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhone2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1081
    Bottom =682
    Left =-1
    Top =-1
    Right =1049
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =294
        Bottom =173
        Top =0
        Name ="Contacts"
        Name =""
    End
End

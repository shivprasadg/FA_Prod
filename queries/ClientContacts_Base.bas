Operation =1
Option =0
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.contactID"
    Alias ="Salutation"
    Expression ="SplitString([contactFullName],\".\",0) & \".\""
    Alias ="FullName"
    Expression ="[contactFirst] & \" \" & IIf(Nz([contactMiddle],\"\")=\"\",Null,[contactMiddle] "
        "& \" \") & [ContactLast]"
    Alias ="Title"
    Expression ="Contacts.contactTitle"
    Alias ="Phone"
    Expression ="Contacts.contactPrimaryPhone"
    Alias ="Mobile"
    Expression ="Contacts.contactCell"
    Alias ="HasEmail"
    Expression ="IIf(Nz([contactEmail],\"\")<>\"\",\"Y\",\"N\")"
    Alias ="EmailAddr"
    Expression ="Contacts.contactEmail"
    Alias ="ClientID"
    Expression ="Contacts.clientRef"
    Expression ="Contacts.contactFullName"
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
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbInteger "ColumnWidth" ="4575"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HasEmail"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salutation"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mobile"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EmailAddr"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =77
    Right =1389
    Bottom =953
    Left =-1
    Top =-1
    Right =1291
    Bottom =548
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =247
        Top =40
        Right =519
        Bottom =430
        Top =0
        Name ="Contacts"
        Name =""
    End
End

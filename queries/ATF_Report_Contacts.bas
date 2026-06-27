Operation =1
Option =0
Begin InputTables
    Name ="ATF_Contacts"
End
Begin OutputColumns
    Expression ="ATF_Contacts.ATFID"
    Expression ="ATF_Contacts.ContactType"
    Expression ="ATF_Contacts.FullName"
    Expression ="ATF_Contacts.EmailYN"
    Alias ="Phone"
    Expression ="ATF_Contacts.contactPrimaryPhone"
    Alias ="Mobile"
    Expression ="ATF_Contacts.contactCell"
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
        dbText "Name" ="ATF_Contacts.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Contacts.ContactType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Contacts.FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Contacts.EmailYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mobile"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1477
    Bottom =1026
    Left =-1
    Top =-1
    Right =1459
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =125
        Top =39
        Right =307
        Bottom =505
        Top =0
        Name ="ATF_Contacts"
        Name =""
    End
End

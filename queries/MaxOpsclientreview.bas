Operation =1
Option =0
Having ="(((Notes.CallType)=\"on-site\") AND ((Contacts.contactClassification)=\"Operatio"
    "ns\"))"
Begin InputTables
    Name ="Notes"
    Name ="Clients"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Notes.CallType"
    Expression ="Contacts.contactClassification"
    Alias ="DT"
    Expression ="Max(IIf(IsNull([noteectdt]),[notedatetime],[noteectdt]))"
End
Begin Joins
    LeftTable ="Notes"
    RightTable ="Clients"
    Expression ="Notes.clientRef = Clients.clientID"
    Flag =1
    LeftTable ="Notes"
    RightTable ="Contacts"
    Expression ="Notes.ContactNM = Contacts.contactFullName"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientID"
    Flag =0
End
Begin Groups
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Notes.CallType"
    GroupLevel =0
    Expression ="Contacts.contactClassification"
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
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.CallType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DT"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1310
    Bottom =793
    Left =-1
    Top =-1
    Right =1278
    Bottom =247
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Notes"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =424
        Top =1
        Right =641
        Bottom =221
        Top =0
        Name ="Contacts"
        Name =""
    End
End

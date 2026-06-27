Operation =1
Option =0
Where ="(((Salesperson.SPInitials)=forms!funotef!fuloex) And ((Notes.noteFU) Is Not Null"
    " And (Notes.noteFU)<=forms!funotef!fupdt))"
Begin InputTables
    Name ="Notes"
    Name ="Clients"
    Name ="Salesperson"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Salesperson.SPInitials"
    Expression ="Notes.noteDateTime"
    Expression ="Notes.noteFU"
    Expression ="Notes.noteMemo"
    Expression ="Contacts.contactFullName"
End
Begin Joins
    LeftTable ="Notes"
    RightTable ="Clients"
    Expression ="Notes.clientRef = Clients.clientID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.SalespersonREF = Salesperson.SalespersonID"
    Flag =1
    LeftTable ="Notes"
    RightTable ="Contacts"
    Expression ="Notes.MAINCONTACT = Contacts.contactID"
    Flag =2
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.noteDateTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.noteFU"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.noteMemo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
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
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =169
        Top =162
        Right =361
        Bottom =306
        Top =0
        Name ="Salesperson"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =765
        Bottom =323
        Top =0
        Name ="Contacts"
        Name =""
    End
End

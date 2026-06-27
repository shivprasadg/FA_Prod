Operation =1
Option =2
Where ="(((Contacts.contactClassification)=\"signator\") AND ((Clients.clienttype)=\"tar"
    "get\")) OR (((Clients.clienttype)=\"target\") AND ((Contacts.[C-Level])=Yes) AND"
    " ((Contacts.contactTitle) Like \"*cfo*\") AND ((Notes.CallType)=\"on-site\")) OR"
    " (((Clients.clienttype)=\"target\") AND ((Contacts.contactTitle) Like \"*Analyst"
    "*\") AND ((Notes.CallType)=\"on-site\")) OR (((Contacts.contactClassification)=\""
    "financial\" Or (Contacts.contactClassification)=\"primary\") AND ((Clients.clien"
    "ttype)=\"target\") AND ((Contacts.[C-Level])=Yes) AND ((Contacts.contactTitle) L"
    "ike \"*cfo*\") AND ((Notes.CallType)=\"phone call\")) OR (((Contacts.contactClas"
    "sification)=\"financial\" Or (Contacts.contactClassification)=\"primary\") AND ("
    "(Clients.clienttype)=\"target\") AND ((Contacts.contactTitle) Like \"*Analyst*\""
    ") AND ((Notes.CallType)=\"phone call\"))"
Begin InputTables
    Name ="Clients"
    Name ="Notes"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Expression ="Contacts.contactID"
    Expression ="Contacts.contactClassification"
    Expression ="Clients.clienttype"
    Expression ="Contacts.[C-Level]"
    Expression ="Contacts.contactTitle"
    Expression ="Notes.CallType"
    Alias ="Sig"
    Expression ="IIf([contactclassification]=\"signator\",0.1,0)"
    Alias ="CCall"
    Expression ="IIf([c-level]=Yes And [calltype]=\"phone call\",0.1,0)"
    Alias ="CVisit"
    Expression ="IIf([c-level]=Yes And [calltype]=\"on-site\",0.2,0)"
    Alias ="ACall"
    Expression ="IIf([c-level]=No And [calltype]=\"phone call\",0.05,0)"
    Alias ="AVisit"
    Expression ="IIf([c-level]=No And [calltype]=\"on-site\",0.1,0)"
End
Begin Joins
    LeftTable ="Notes"
    RightTable ="Contacts"
    Expression ="Notes.MainContact = Contacts.contactID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Notes"
    Expression ="Clients.clientID = Notes.clientRef"
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
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clienttype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.[C-Level]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes.CallType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ACall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AVisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1095
    Bottom =631
    Left =-1
    Top =-1
    Right =1063
    Bottom =101
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =235
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =343
        Top =14
        Right =484
        Bottom =144
        Top =0
        Name ="Notes"
        Name =""
    End
    Begin
        Left =528
        Top =12
        Right =736
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
End

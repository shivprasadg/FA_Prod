Operation =1
Option =0
Where ="(((Contacts.ClientREF)=Forms!Units!clientID) And ((Contacts.contactClassificatio"
    "n)=\"Location\"))"
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.contactID"
    Expression ="Contacts.LocationName"
    Alias ="address"
    Expression ="[Contacts]![contactAddress1]+\" \"+[Contacts]![contactCity]+\", \"+[Contacts]![L"
        "State]"
    Expression ="Contacts.ClientREF"
    Expression ="Contacts.contactClassification"
End
Begin OrderBy
    Expression ="Contacts.LocationName"
    Flag =0
    Expression ="[Contacts]![contactAddress1]+\" \"+[Contacts]![contactCity]+\", \"+[Contacts]![L"
        "State]"
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
        dbText "Name" ="address"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
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
    Bottom =190
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
End

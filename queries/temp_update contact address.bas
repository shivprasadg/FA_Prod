Operation =4
Option =0
Where ="(((Contacts.clientRef)=Forms!Contacts_view!clientRef) And ((Contacts.contactID)="
    "Forms!Contacts_view!contactID))"
Begin InputTables
    Name ="Clients"
    Name ="Contacts"
End
Begin OutputColumns
    Name ="Contacts.contactAddress1"
    Expression ="[clients].[clientAddress1]"
    Name ="Contacts.contactAddress2"
    Expression ="[clients].[clientAddress2]"
    Name ="Contacts.contactCity"
    Expression ="[clients].[clientcity]"
    Name ="Contacts.LState"
    Expression ="[clients].[clientstate]"
    Name ="Contacts.contactZip"
    Expression ="[clients].[clientzip]"
    Name ="Contacts.contactCountry"
    Expression ="[clients].[clientcountry]"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Contacts"
    Expression ="Clients.clientID = Contacts.clientRef"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Contacts.contactCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.clientRef"
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
    Bottom =174
    Left =0
    Top =0
    ColumnsShown =579
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
End

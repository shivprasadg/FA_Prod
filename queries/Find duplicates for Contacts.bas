Operation =1
Option =0
Where ="(((Contacts.contactID) In (SELECT [contactID] FROM [Contacts] As Tmp GROUP BY [c"
    "ontactID] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Contacts.contactID"
    Expression ="Contacts.clientRef"
End
Begin OrderBy
    Expression ="Contacts.contactID"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1322
    Bottom =793
    Left =-1
    Top =-1
    Right =1290
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Contacts"
        Name =""
    End
End

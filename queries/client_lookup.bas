Operation =1
Option =0
Begin InputTables
    Name ="clients"
End
Begin OutputColumns
    Expression ="clients.clientID"
    Alias ="CompanyAndDBA"
    Expression ="[clientcompanyName] & IIf([clientDBA]<>\"\",\" d/b/a \" & [clientDBA],\"\")"
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
        dbText "Name" ="CompanyAndDBA"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =147
    Right =1054
    Bottom =594
    Left =-1
    Top =-1
    Right =975
    Bottom =235
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =242
        Bottom =242
        Top =0
        Name ="clients"
        Name =""
    End
End

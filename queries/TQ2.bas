Operation =1
Option =0
Begin InputTables
    Name ="TQ1"
    Alias ="CTT"
End
Begin OutputColumns
    Alias ="FA-Asset Num"
    Expression ="CTT.[FA TitleLog Num]"
    Alias ="Expr1"
    Expression ="CTT.VIN"
    Alias ="Expr2"
    Expression ="CTT.[File Type]"
    Alias ="Expr3"
    Expression ="CTT.[Portfolio Certificate No]"
    Alias ="Expr4"
    Expression ="CTT.[Lender or Holder Name]"
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
    Right =1268
    Bottom =1271
    Left =-1
    Top =-1
    Right =1244
    Bottom =951
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CTT"
        Name =""
    End
End

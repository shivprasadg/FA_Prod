Operation =1
Option =0
Where ="(((PROMNT.PNExp)>Date()))"
Begin InputTables
    Name ="Lines"
    Name ="PROMNT"
End
Begin OutputColumns
    Expression ="Lines.Lineid"
    Alias ="SumOfPNAmt"
    Expression ="Sum(PROMNT.PNAmt)"
End
Begin Joins
    LeftTable ="Lines"
    RightTable ="PROMNT"
    Expression ="Lines.Lineid = PROMNT.LineID"
    Flag =1
End
Begin OrderBy
    Expression ="Lines.Lineid"
    Flag =0
End
Begin Groups
    Expression ="Lines.Lineid"
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
        dbText "Name" ="Lines.Lineid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPNAmt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =888
    Bottom =702
    Left =-1
    Top =-1
    Right =856
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Lines"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =584
        Bottom =272
        Top =0
        Name ="PROMNT"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((Rmkt.[UnitRef]) In (SELECT [UnitRef] FROM [Rmkt] As Tmp GROUP BY [UnitRef] HA"
    "VING Count(*)>1 )))"
Begin InputTables
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Rmkt.[UnitRef]"
    Expression ="Rmkt.[SUF Sent]"
    Expression ="Rmkt.[SUF Returned]"
    Expression ="Rmkt.[BuyOutDate]"
    Expression ="Rmkt.[BuyOutCost]"
    Expression ="Rmkt.[BuyerRef]"
End
Begin OrderBy
    Expression ="Rmkt.[UnitRef]"
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
    Right =1495
    Bottom =1038
    Left =-1
    Top =-1
    Right =1463
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
End

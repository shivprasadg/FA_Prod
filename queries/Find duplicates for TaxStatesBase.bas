Operation =1
Option =0
Where ="(((TaxStatesBase.[TaxTypeId]) In (SELECT [TaxTypeId] FROM [TaxStatesBase] As Tmp"
    " GROUP BY [TaxTypeId] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="TaxStatesBase"
End
Begin OutputColumns
    Expression ="TaxStatesBase.[TaxTypeId]"
    Expression ="TaxStatesBase.[TaxStatesID]"
    Expression ="TaxStatesBase.[StartsOn]"
    Expression ="TaxStatesBase.[EndsOn]"
End
Begin OrderBy
    Expression ="TaxStatesBase.[TaxTypeId]"
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
    Begin
        dbText "Name" ="[TaxTypeId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[TaxStatesID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[StartsOn]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[EndsOn]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1390
    Bottom =969
    Left =-1
    Top =-1
    Right =1328
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
        Name ="TaxStatesBase"
        Name =""
    End
End

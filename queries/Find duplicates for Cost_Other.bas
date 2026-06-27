Operation =1
Option =0
Where ="(((Cost_Other.AssetId) In (SELECT [AssetId] FROM [Cost_Other] As Tmp GROUP BY [A"
    "ssetId] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="Cost_Other"
End
Begin OutputColumns
    Expression ="Cost_Other.AssetId"
    Expression ="Cost_Other.OtherMake"
    Expression ="Cost_Other.ChildType"
    Expression ="Cost_Other.OtherInvoice"
End
Begin OrderBy
    Expression ="Cost_Other.AssetId"
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
    Right =1591
    Bottom =872
    Left =-1
    Top =-1
    Right =1559
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
        Name ="Cost_Other"
        Name =""
    End
End

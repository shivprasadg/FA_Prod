Operation =1
Option =0
Where ="(((LeaseReport2.AssetId) In (SELECT [AssetId] FROM [LeaseReport2] As Tmp GROUP B"
    "Y [AssetId] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="LeaseReport2"
End
Begin OutputColumns
    Expression ="LeaseReport2.AssetId"
    Expression ="LeaseReport2.FAID"
    Expression ="LeaseReport2.UnitType"
    Expression ="LeaseReport2.Make"
End
Begin OrderBy
    Expression ="LeaseReport2.AssetId"
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
    Right =1308
    Bottom =1038
    Left =-1
    Top =-1
    Right =1276
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="LeaseReport2"
        Name =""
    End
End

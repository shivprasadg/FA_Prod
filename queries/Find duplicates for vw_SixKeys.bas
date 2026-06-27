Operation =1
Option =0
Where ="(((vw_SixKeys.VIN) In (SELECT [VIN] FROM [vw_SixKeys] As Tmp GROUP BY [VIN] HAVI"
    "NG Count(*)>1 )))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.VIN"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.SGrpID"
    Expression ="vw_SixKeys.CompanyName"
End
Begin OrderBy
    Expression ="vw_SixKeys.VIN"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[Find duplicates for vw_SixKeys].[VIN], [Find duplicates for vw_SixKeys].[Client"
    "GroupId]"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =87
    Right =1876
    Bottom =1021
    Left =-1
    Top =-1
    Right =1621
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End

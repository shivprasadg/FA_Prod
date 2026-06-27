Operation =1
Option =1
Where ="SchID=2958"
Begin InputTables
    Name ="ATF_Report_ChildPartsActive"
End
Begin OutputColumns
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
        dbText "Name" ="ATF_Report_ChildPartsActive.SpecProfile"
        dbInteger "ColumnWidth" ="6450"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_ChildPartsActive.ChildDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_ChildPartsActive.Child Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_ChildPartsActive.PartCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =79
    Top =210
    Right =1267
    Bottom =1052
    Left =-1
    Top =-1
    Right =1170
    Bottom =412
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ATF_Report_ChildPartsActive"
        Name =""
    End
End

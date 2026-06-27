Operation =1
Option =1
Where ="SchID=2958"
Begin InputTables
    Name ="ATF_Report_SchNotesActive"
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
        dbText "Name" ="ATF_Report_SchNotesActive.NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.FieldGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.ActionDate"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.Sorter"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.ActionDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.SysUser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.RevID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.ReviewedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Report_SchNotesActive.ReviewedDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =307
    Top =21
    Right =1121
    Bottom =603
    Left =-1
    Top =-1
    Right =790
    Bottom =258
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ATF_Report_SchNotesActive"
        Name =""
    End
End

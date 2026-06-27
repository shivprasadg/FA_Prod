Operation =1
Option =0
Where ="(((UnitReportStartEndDates.ReportEndDate) Is Null Or (UnitReportStartEndDates.Re"
    "portEndDate)<>[UnitReportStartEndDates].[EndDate]))"
Begin InputTables
    Name ="UnitReportStartEndDates"
End
Begin OutputColumns
    Expression ="UnitReportStartEndDates.UnitID"
    Expression ="UnitReportStartEndDates.ReportStartDate"
    Expression ="UnitReportStartEndDates.ReportEndDate"
    Expression ="UnitReportStartEndDates.StartDate"
    Expression ="UnitReportStartEndDates.EndDate"
    Expression ="UnitReportStartEndDates.UnitStatus"
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
        dbText "Name" ="UnitReportStartEndDates.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitReportStartEndDates.ReportStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitReportStartEndDates.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitReportStartEndDates.StartDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitReportStartEndDates.EndDate"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitReportStartEndDates.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =187
    Top =14
    Right =1300
    Bottom =942
    Left =-1
    Top =-1
    Right =1081
    Bottom =542
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =224
        Top =105
        Right =471
        Bottom =402
        Top =0
        Name ="UnitReportStartEndDates"
        Name =""
    End
End

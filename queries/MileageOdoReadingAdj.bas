Operation =1
Option =0
Where ="(((vw_UMP_Mileage.OBCDataStatus)<>6))"
Begin InputTables
    Name ="vw_UMP_Mileage"
End
Begin OutputColumns
    Expression ="vw_UMP_Mileage.UnitID"
    Alias ="FirstReport"
    Expression ="Min(vw_UMP_Mileage.LogActivityDate)"
    Alias ="LastReport"
    Expression ="Max(vw_UMP_Mileage.LogActivityDate)"
    Alias ="MthsRptd"
    Expression ="DateDiff('m',Min([LogActivityDate]),Max([LogActivityDate]))"
    Alias ="ODOMiles"
    Expression ="Round(Max([OdometerEnd])-Min([OdometerBegin]),1)"
    Alias ="MPM"
    Expression ="IIf(Nz([ODOMiles],0)>0 And Nz([MthsRptd],0)>0,Round([ODOMiles]/[MthsRptd],1),0)"
End
Begin Groups
    Expression ="vw_UMP_Mileage.UnitID"
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
        dbText "Name" ="LastReport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstReport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODOMiles"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthReported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MPM"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MthsRpted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MthsRptd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =69
    Top =171
    Right =1336
    Bottom =720
    Left =-1
    Top =-1
    Right =1235
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =146
        Top =35
        Right =495
        Bottom =329
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
End

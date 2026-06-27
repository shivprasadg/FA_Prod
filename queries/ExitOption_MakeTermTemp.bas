Operation =2
Name ="ExitTermTemp"
Option =8
Begin InputTables
    Name ="EXITOptionsGAPReport"
    Name ="EXITOptions_WorkTable"
End
Begin OutputColumns
    Expression ="EXITOptionsGAPReport.AssetID"
    Alias ="OptionTypeID"
    Expression ="2"
    Alias ="TerminateEarly"
    Expression ="IIf([Early]=\"Yes\",-1,0)"
    Alias ="Residual"
    Expression ="EXITOptionsGAPReport.ProjRV"
    Alias ="TerminateFee"
    Expression ="EXITOptionsGAPReport.ProjTermPenalty"
    Alias ="Notes"
    Expression ="EXITOptionsGAPReport.note"
    Alias ="EndOrActualDate"
    Expression ="EXITOptionsGAPReport.OffLeaseDate"
End
Begin Joins
    LeftTable ="EXITOptions_WorkTable"
    RightTable ="EXITOptionsGAPReport"
    Expression ="EXITOptions_WorkTable.UnitID = EXITOptionsGAPReport.AssetID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="EXITOptionsGAPReport.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.ProjRV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TerminateEarly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbInteger "ColumnWidth" ="5520"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.GAP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.MonthsDiff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.*"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateEarly"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.Early"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.ProjTermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TerminateFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =67
    Top =-5
    Right =1616
    Bottom =946
    Left =-1
    Top =-1
    Right =1517
    Bottom =616
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =190
        Top =81
        Right =334
        Bottom =357
        Top =0
        Name ="EXITOptionsGAPReport"
        Name =""
    End
    Begin
        Left =6
        Top =51
        Right =150
        Bottom =195
        Top =0
        Name ="EXITOptions_WorkTable"
        Name =""
    End
End

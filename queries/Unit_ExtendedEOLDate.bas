Operation =1
Option =0
Having ="(((UnitExitOption.OptionTypeID)=3))"
Begin InputTables
    Name ="UnitExitOption"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
    Expression ="UnitExitOption.StartOrPurposedDate"
    Expression ="UnitExitOption.OptionTypeID"
    Alias ="EndOrActualDate"
    Expression ="Max(UnitExitOption.EndOrActualDate)"
    Alias ="TotalMonths"
    Expression ="Sum(UnitExitOption.TermMonths)"
    Alias ="TimesExtended"
    Expression ="Count(UnitExitOption.UnitId)"
    Expression ="UnitExitOption.Fees"
End
Begin Groups
    Expression ="UnitExitOption.UnitID"
    GroupLevel =0
    Expression ="UnitExitOption.StartOrPurposedDate"
    GroupLevel =0
    Expression ="UnitExitOption.OptionTypeID"
    GroupLevel =0
    Expression ="UnitExitOption.Fees"
    GroupLevel =0
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
dbMemo "OrderBy" ="[Unit_ExtendedEOLDate].[StartOrPurposedDate] DESC"
Begin
    Begin
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TimesExtended"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitExitOption.Fees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.StartOrPurposedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extended"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =159
    Top =42
    Right =1451
    Bottom =918
    Left =-1
    Top =-1
    Right =1274
    Bottom =548
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =77
        Top =85
        Right =325
        Bottom =571
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
End

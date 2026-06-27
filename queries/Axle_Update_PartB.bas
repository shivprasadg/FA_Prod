Operation =1
Option =0
Having ="((([UnitVinDecoder].[DRIVE_LINE_TYPE])<>\"\"))"
Begin InputTables
    Name ="UnitVinDecoder"
End
Begin OutputColumns
    Alias ="VinPreFix_Decoder"
    Expression ="Left([VIN],11)"
    Alias ="Expr1"
    Expression ="UnitVinDecoder.DRIVE_LINE_TYPE"
    Alias ="NewAxel"
    Expression ="IIf([DRIVE_LINE_TYPE]=\"6x4\",\"T/A\",IIf([DRIVE_LINE_TYPE]=\"4x2\",\"S/A\",[DRI"
        "VE_LINE_TYPE]))"
    Expression ="UnitVinDecoder.MAKE"
    Expression ="UnitVinDecoder.MODEL"
End
Begin Groups
    Expression ="Left([VIN],11)"
    GroupLevel =0
    Expression ="UnitVinDecoder.DRIVE_LINE_TYPE"
    GroupLevel =0
    Expression ="IIf([DRIVE_LINE_TYPE]=\"6x4\",\"T/A\",IIf([DRIVE_LINE_TYPE]=\"4x2\",\"S/A\",[DRI"
        "VE_LINE_TYPE]))"
    GroupLevel =0
    Expression ="UnitVinDecoder.MAKE"
    GroupLevel =0
    Expression ="UnitVinDecoder.MODEL"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="UnitVinDecoder.DRIVE_LINE_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewAxel"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VinPreFix_Decoder"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.MODEL"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1520
    Bottom =825
    Left =-1
    Top =-1
    Right =1488
    Bottom =451
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =364
        Top =90
        Right =626
        Bottom =475
        Top =0
        Name ="UnitVinDecoder"
        Name =""
    End
End

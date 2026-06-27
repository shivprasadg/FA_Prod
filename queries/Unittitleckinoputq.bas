Operation =1
Option =0
Where ="(((Units.unitvin)=[forms]![titlecheckin]![seekvin]))"
Begin InputTables
    Name ="Units"
    Name ="UnitTitleLog"
    Name ="Reasons"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="UnitTitleLog.UnitTitleLogDate"
    Alias ="Status"
    Expression ="IIf([entrytypeid]=1,\"In\",\"Out\")"
    Expression ="Reasons.ReasonDescription"
    Expression ="UnitTitleLog.TitleNumber"
    Expression ="Units.unitvin"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="UnitTitleLog"
    Expression ="Units.UnitID = UnitTitleLog.UnitRef"
    Flag =1
    LeftTable ="UnitTitleLog"
    RightTable ="Reasons"
    Expression ="UnitTitleLog.EntryReasonId = Reasons.ReasonId"
    Flag =1
End
Begin OrderBy
    Expression ="UnitTitleLog.UnitTitleLogDate"
    Flag =1
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
        dbText "Name" ="UnitTitleLog.UnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTitleLog.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reasons.ReasonDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1107
    Bottom =662
    Left =-1
    Top =-1
    Right =1075
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =748
        Bottom =156
        Top =0
        Name ="UnitTitleLog"
        Name =""
    End
    Begin
        Left =796
        Top =12
        Right =940
        Bottom =156
        Top =0
        Name ="Reasons"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((vw_SixKeys.mlorig)=\"FA\"))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="BuildYear"
    Expression ="DatePart(\"yyyy\",[UnitBuildDate])"
    Expression ="Units.UnitStatus"
    Alias ="Active"
    Expression ="IIf([UnitStatus]=\"A\",Count([UnitStatus]),0)"
    Alias ="BeingBuilt"
    Expression ="IIf([UnitStatus]=\"B\",Count([UnitStatus]),0)"
    Alias ="Casualty"
    Expression ="IIf([UnitStatus]=\"C\",Count([UnitStatus]),0)"
    Alias ="Brokered"
    Expression ="IIf([UnitStatus]=\"K\",Count([UnitStatus]),0)"
    Alias ="SoldNotOwn"
    Expression ="IIf([UnitStatus]=\"KS\",Count([UnitStatus]),0)"
    Alias ="OffLease"
    Expression ="IIf([UnitStatus]=\"O\",Count([UnitStatus]),0)"
    Alias ="Remarketing"
    Expression ="IIf([UnitStatus]=\"R\",Count([UnitStatus]),0)"
    Alias ="SoldOwn"
    Expression ="IIf([UnitStatus]=\"S\",Count([UnitStatus]),0)"
    Alias ="Tracking"
    Expression ="IIf([UnitStatus]=\"T\",Count([UnitStatus]),0)"
    Alias ="TrackPending"
    Expression ="IIf([UnitStatus]=\"TR\",Count([UnitStatus]),0)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="DatePart(\"yyyy\",[UnitBuildDate])"
    Flag =0
End
Begin Groups
    Expression ="DatePart(\"yyyy\",[UnitBuildDate])"
    GroupLevel =0
    Expression ="Units.UnitStatus"
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
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BeingBuilt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Casualty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Brokered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remarketing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldOwn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldNotOwn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracking"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrackPending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.mlorig"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =165
    Top =39
    Right =1547
    Bottom =713
    Left =-1
    Top =-1
    Right =1364
    Bottom =404
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =111
        Top =-11
        Right =521
        Bottom =432
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =569
        Top =12
        Right =713
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End

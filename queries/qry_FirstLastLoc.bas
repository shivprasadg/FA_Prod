Operation =1
Option =0
Having ="(((LocChangeT.unitref)=[forms]![UnitsForm]![unitid]))"
Begin InputTables
    Name ="LocChangeT"
End
Begin OutputColumns
    Expression ="LocChangeT.ID"
    Alias ="FirstCHdt1"
    Expression ="First(LocChangeT.LocChDt)"
    Alias ="LastCHdt2"
    Expression ="First(LocChangeT.LocChDt)"
    Alias ="FirstOfOldUnitloc"
    Expression ="First(LocChangeT.OldUnitloc)"
    Alias ="FirstOldStreet"
    Expression ="First(LocChangeT.OLocationStreet)"
    Alias ="FirstOldCity"
    Expression ="First(LocChangeT.OCity)"
    Alias ="FirstOldCounty"
    Expression ="First(LocChangeT.OCounty)"
    Alias ="FirstOldState"
    Expression ="First(LocChangeT.OSt)"
    Alias ="FirstOldZip"
    Expression ="First(LocChangeT.Ozip)"
    Alias ="FirstOldTaxStatus"
    Expression ="First(LocChangeT.OldTxstatus)"
    Alias ="FirstoldTaxNote"
    Expression ="First(LocChangeT.OTaxNotes)"
    Alias ="LastOfUnitsFormloc"
    Expression ="Last(LocChangeT.UnitsFormloc)"
    Alias ="LastNewStreet"
    Expression ="Last(LocChangeT.NLocationStreet)"
    Alias ="LastNewTaxStatus"
    Expression ="Last(LocChangeT.Newtaxstatus)"
    Alias ="LastNewTaxNote"
    Expression ="Last(LocChangeT.NewTaxnotes)"
    Alias ="NewLocation"
    Expression ="Last(LocChangeT.NCity)"
    Alias ="LastNewCounty"
    Expression ="Last(LocChangeT.NCounty)"
    Alias ="LastNewState"
    Expression ="Last(LocChangeT.NSt)"
    Alias ="LastNewZip"
    Expression ="Last(LocChangeT.Nzip)"
    Alias ="MinRptChDate"
    Expression ="First(LocChangeT.LocRptDt)"
    Alias ="MaxRptChDate"
    Expression ="Last(LocChangeT.LocRptDt)"
    Expression ="LocChangeT.unitref"
End
Begin Groups
    Expression ="LocChangeT.ID"
    GroupLevel =0
    Expression ="LocChangeT.unitref"
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
        dbText "Name" ="FirstOldCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOldCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstoldTaxNote"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastNewStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldStreet"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOldState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MinRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstCHdt1"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastCHdt2"
        dbInteger "ColumnWidth" ="3015"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewCounty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastNewZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfOldUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfNewUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocChangeT.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocChangeT.unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfUnitsFormloc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =57
    Right =1824
    Bottom =1066
    Left =-1
    Top =-1
    Right =1711
    Bottom =435
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =63
        Top =-8
        Right =312
        Bottom =477
        Top =0
        Name ="LocChangeT"
        Name =""
    End
End

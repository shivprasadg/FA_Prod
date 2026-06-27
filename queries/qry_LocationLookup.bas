Operation =1
Option =0
Where ="(((LocChangeT.unitref)=[forms]![UnitsForm]![unitid]))"
Begin InputTables
    Name ="LocChangeT"
End
Begin OutputColumns
    Expression ="LocChangeT.ID"
    Alias ="FirstCHdt1"
    Expression ="LocChangeT.LocChDt"
    Alias ="LastCHdt2"
    Expression ="LocChangeT.LocChDt"
    Expression ="LocChangeT.OldUnitloc"
    Alias ="FirstOldStreet"
    Expression ="LocChangeT.OLocationStreet"
    Alias ="FirstOldCity"
    Expression ="LocChangeT.OCity"
    Alias ="FirstOldCounty"
    Expression ="LocChangeT.OCounty"
    Alias ="FirstOldState"
    Expression ="LocChangeT.OSt"
    Alias ="FirstOldZip"
    Expression ="LocChangeT.Ozip"
    Alias ="FirstOldTaxStatus"
    Expression ="LocChangeT.OldTxstatus"
    Alias ="FirstoldTaxNote"
    Expression ="LocChangeT.OTaxNotes"
    Alias ="LastNewStreet"
    Expression ="LocChangeT.NLocationStreet"
    Alias ="LastNewTaxStatus"
    Expression ="LocChangeT.Newtaxstatus"
    Alias ="LastNewTaxNote"
    Expression ="LocChangeT.NewTaxnotes"
    Alias ="NewLocation"
    Expression ="LocChangeT.NCity"
    Alias ="LastNewCounty"
    Expression ="LocChangeT.NCounty"
    Alias ="LastNewState"
    Expression ="LocChangeT.NSt"
    Alias ="LastNewZip"
    Expression ="LocChangeT.Nzip"
    Alias ="MinRptChDate"
    Expression ="LocChangeT.LocRptDt"
    Alias ="MaxRptChDate"
    Expression ="LocChangeT.LocRptDt"
    Expression ="LocChangeT.unitref"
End
Begin OrderBy
    Expression ="LocChangeT.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="2"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="LocChangeT.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstCHdt1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastCHdt2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOldStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOldTaxStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Yes/No"
    End
    Begin
        dbText "Name" ="FirstoldTaxNote"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3480"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastNewStreet"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LastNewTaxStatus"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="Yes/No"
    End
    Begin
        dbText "Name" ="LastNewTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNewZip"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="LocChangeT.unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocChangeT.OldUnitloc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =148
    Top =175
    Right =1406
    Bottom =821
    Left =-1
    Top =-1
    Right =1226
    Bottom =425
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =277
        Bottom =433
        Top =0
        Name ="LocChangeT"
        Name =""
    End
End

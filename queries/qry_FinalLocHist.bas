Operation =1
Option =0
Having ="(((qry_FirstLastLoc.unitref)=2359))"
Begin InputTables
    Name ="qry_FirstLastLoc"
End
Begin OutputColumns
    Alias ="FirstOfFirstCHdt1"
    Expression ="First(qry_FirstLastLoc.FirstCHdt1)"
    Alias ="LastOfLastCHdt2"
    Expression ="Last(qry_FirstLastLoc.LastCHdt2)"
    Alias ="FirstOfFirstOfOldUnitloc"
    Expression ="First(qry_FirstLastLoc.FirstOfOldUnitloc)"
    Alias ="FirstOfFirstOldStreet"
    Expression ="First(qry_FirstLastLoc.FirstOldStreet)"
    Alias ="FirstOfFirstOldCity"
    Expression ="First(qry_FirstLastLoc.FirstOldCity)"
    Alias ="FirstOfFirstOldCounty"
    Expression ="First(qry_FirstLastLoc.FirstOldCounty)"
    Alias ="FirstOfFirstOldState"
    Expression ="First(qry_FirstLastLoc.FirstOldState)"
    Alias ="FirstOfFirstOldZip"
    Expression ="First(qry_FirstLastLoc.FirstOldZip)"
    Alias ="FirstOfFirstOldTaxStatus"
    Expression ="First(qry_FirstLastLoc.FirstOldTaxStatus)"
    Alias ="FirstOfFirstoldTaxNote"
    Expression ="First(qry_FirstLastLoc.FirstoldTaxNote)"
    Alias ="LastOfLastOfUnitsFormloc"
    Expression ="Last(qry_FirstLastLoc.LastOfUnitsFormloc)"
    Alias ="LastOfLastNewStreet"
    Expression ="Last(qry_FirstLastLoc.LastNewStreet)"
    Alias ="LastOfLastNewTaxStatus"
    Expression ="Last(qry_FirstLastLoc.LastNewTaxStatus)"
    Alias ="LastOfLastNewTaxNote"
    Expression ="Last(qry_FirstLastLoc.LastNewTaxNote)"
    Alias ="LastOfNewLocation"
    Expression ="Last(qry_FirstLastLoc.NewLocation)"
    Alias ="LastOfLastNewCounty"
    Expression ="Last(qry_FirstLastLoc.LastNewCounty)"
    Alias ="LastOfLastNewState"
    Expression ="Last(qry_FirstLastLoc.LastNewState)"
    Alias ="LastOfLastNewZip"
    Expression ="Last(qry_FirstLastLoc.LastNewZip)"
    Alias ="LastOfMinRptChDate"
    Expression ="Last(qry_FirstLastLoc.MinRptChDate)"
    Alias ="LastOfMaxRptChDate"
    Expression ="Last(qry_FirstLastLoc.MaxRptChDate)"
    Expression ="qry_FirstLastLoc.unitref"
End
Begin Groups
    Expression ="qry_FirstLastLoc.unitref"
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
        dbText "Name" ="FirstOfFirstCHdt1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOfFirstOfOldUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstOldTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfFirstoldTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastOfNewUnitloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewTaxStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewTaxNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfNewLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastNewZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfMinRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfMaxRptChDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastCHdt2"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfLastOfUnitsFormloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_FirstLastLoc.unitref"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =162
    Top =59
    Right =1832
    Bottom =874
    Left =-1
    Top =-1
    Right =1638
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =237
        Bottom =471
        Top =0
        Name ="qry_FirstLastLoc"
        Name =""
    End
End

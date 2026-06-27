Operation =1
Option =0
Where ="(((UMP.[INDX]) In (SELECT        MAX(INDX) \015\012                             "
    "  FROM            vw_UnitMileageProcessed AS UMP2\015\012                       "
    "        GROUP BY UnitID\015\012\015\012)))"
Begin InputTables
    Name ="vw_UnitMileageProcessed"
    Alias ="UMP"
End
Begin OutputColumns
    Expression ="UMP.UnitID"
    Alias ="CurrentDate"
    Expression ="UMP.LogActivityDate"
    Alias ="CurrentMiles"
    Expression ="UMP.Miles"
    Alias ="CurrentFuel"
    Expression ="UMP.Fuel"
    Alias ="CurrentFuelMov"
    Expression ="UMP.FuelMoving"
    Alias ="CurrentHours"
    Expression ="UMP.Hours"
    Expression ="UMP.INDX"
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
        dbText "Name" ="CurrentMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrentFuel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrentHours"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UMP.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrentFuelMov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UMP.INDX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1561
    Bottom =984
    Left =-1
    Top =-1
    Right =1529
    Bottom =626
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =363
        Bottom =382
        Top =0
        Name ="UMP"
        Name =""
    End
End

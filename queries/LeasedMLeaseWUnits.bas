Operation =1
Option =0
Where ="(((Leased_MLSchGrpUnits.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Leased_MLSchGrpUnits"
    Name ="LeasedClientsWUnits"
End
Begin OutputColumns
    Expression ="Leased_MLSchGrpUnits.MLID"
    Expression ="Leased_MLSchGrpUnits.MLNo"
    Expression ="Leased_MLSchGrpUnits.MLOrig"
    Expression ="Leased_MLSchGrpUnits.Client"
End
Begin Joins
    LeftTable ="LeasedClientsWUnits"
    RightTable ="Leased_MLSchGrpUnits"
    Expression ="LeasedClientsWUnits.clientID = Leased_MLSchGrpUnits.Client"
    Flag =1
End
Begin OrderBy
    Expression ="Leased_MLSchGrpUnits.MLNo"
    Flag =0
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
        dbText "Name" ="Leased_MLSchGrpUnits.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leased_MLSchGrpUnits.MLNo"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leased_MLSchGrpUnits.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leased_MLSchGrpUnits.Client"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =10
    Top =79
    Right =1503
    Bottom =862
    Left =-1
    Top =-1
    Right =1461
    Bottom =364
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =444
        Top =12
        Right =588
        Bottom =275
        Top =0
        Name ="Leased_MLSchGrpUnits"
        Name =""
    End
    Begin
        Left =203
        Top =15
        Right =347
        Bottom =159
        Top =0
        Name ="LeasedClientsWUnits"
        Name =""
    End
End

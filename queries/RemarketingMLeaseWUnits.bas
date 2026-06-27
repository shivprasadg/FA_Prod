Operation =1
Option =0
Begin InputTables
    Name ="RemarketingClientsWUnits"
    Name ="ReMarketing_MLSchGrpUnits"
End
Begin OutputColumns
    Expression ="ReMarketing_MLSchGrpUnits.MLID"
    Expression ="ReMarketing_MLSchGrpUnits.MLNo"
    Expression ="ReMarketing_MLSchGrpUnits.MLOrig"
    Expression ="ReMarketing_MLSchGrpUnits.Client"
End
Begin Joins
    LeftTable ="RemarketingClientsWUnits"
    RightTable ="ReMarketing_MLSchGrpUnits"
    Expression ="RemarketingClientsWUnits.ClientID = ReMarketing_MLSchGrpUnits.Client"
    Flag =1
End
Begin OrderBy
    Expression ="ReMarketing_MLSchGrpUnits.MLNo"
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
        dbText "Name" ="ReMarketing_MLSchGrpUnits.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_MLSchGrpUnits.MLNo"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_MLSchGrpUnits.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReMarketing_MLSchGrpUnits.Client"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1493
    Bottom =823
    Left =-1
    Top =-1
    Right =1469
    Bottom =449
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =69
        Top =55
        Right =240
        Bottom =199
        Top =0
        Name ="RemarketingClientsWUnits"
        Name =""
    End
    Begin
        Left =282
        Top =131
        Right =496
        Bottom =353
        Top =0
        Name ="ReMarketing_MLSchGrpUnits"
        Name =""
    End
End

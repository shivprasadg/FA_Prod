Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.PONum1"
    Expression ="SchGrp.POVendor"
    Expression ="SchGrp.POSentDt"
    Expression ="SchGrp.PONum2"
    Expression ="SchGrp.PO2Vendor"
    Expression ="SchGrp.PO2SentDt"
    Expression ="SchGrp.PONum3"
    Expression ="SchGrp.PO3Vendor"
    Expression ="SchGrp.PO3SentDt"
    Expression ="SchGrp.PONum4"
    Expression ="SchGrp.PO4Vendor"
    Expression ="SchGrp.PO4SentDt"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([GroupPOs].[SGrpID]=2133)"
Begin
    Begin
        dbText "Name" ="SchGrp.PO3SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO4SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO2SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO2Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO3Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO4Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =127
    Top =96
    Right =1614
    Bottom =820
    Left =-1
    Top =-1
    Right =1469
    Bottom =641
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End

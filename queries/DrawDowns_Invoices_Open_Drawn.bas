Operation =1
Option =0
Begin InputTables
    Name ="vw_DrawDowns_Invoices"
    Alias ="DDI"
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
            dbText "Name" ="DDI.Invoice"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="Open"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DDI.sGrpID"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DDI.Vendor"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="Open Amt"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="SchGrp"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="Drawn Amt"
            dbLong "AggregateType" ="-1"
            dbInteger "ColumnWidth" ="2415"
            dbBoolean "ColumnHidden" ="0"
        End
        Begin
            dbText "Name" ="DDI.VendorID"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DDI.ClientID"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DDI.SchID"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="Drawn"
            dbInteger "ColumnWidth" ="1020"
            dbBoolean "ColumnHidden" ="0"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DDI.MLID"
            dbLong "AggregateType" ="-1"
        End
        Begin
            dbText "Name" ="DrawDownId"
            dbLong "AggregateType" ="-1"
            dbInteger "ColumnWidth" ="1620"
            dbBoolean "ColumnHidden" ="0"
        End
        Begin
            dbText "Name" ="IsSplit.Split"
            dbLong "AggregateType" ="-1"
        End
    End
    Begin
        State =0
        Left =16
        Top =190
        Right =1587
        Bottom =1080
        Left =-1
        Top =-1
        Right =1547
        Bottom =423
        Left =0
        Top =0
        ColumnsShown =543
        Begin
            Left =48
            Top =12
            Right =192
            Bottom =464
            Top =0
            Name ="DDI"
            Name =""
        End
        Begin
            Left =231
            Top =13
            Right =375
            Bottom =157
            Top =0
            Name ="IsSplit"
            Name =""
        End
    End

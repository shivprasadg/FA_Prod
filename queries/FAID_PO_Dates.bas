Operation =1
Option =0
Where ="(((Clients.ClientShNm) Not Like \"%Dem%\") AND ((Schedule.SchFAID) Like \"14%\" "
    "Or (Schedule.SchFAID) Like \"13%\" Or (Schedule.SchFAID) Like \"12%\" Or (Schedu"
    "le.SchFAID) Like \"11%\" Or (Schedule.SchFAID) Like \"10%\") AND ((MstrLease.MLO"
    "rig)=\"FA\") AND ((Schedule.SchID) Is Not Null) AND ((SchGrp.SGrpID) Is Not Null"
    "))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Vendors"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Alias ="FAID"
    Expression ="Schedule.SchFAID"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="POnumber"
    Expression ="SchGrp.PONum1"
    Alias ="POSent"
    Expression ="SchGrp.POSentDt"
    Expression ="Vendors.VendorName"
    Alias ="EarlyEstimated"
    Expression ="CDate(DMin(\"UnitEstDelDate\",\"Units\",\"GroupREF=\" & [SchGrp].[SGrpID] & \"\""
        "))"
    Alias ="EarlyBuild"
    Expression ="CDate(DMin(\"UnitBuildDate\",\"Units\",\"GroupREF=\" & [SchGrp].[SGrpID] & \"\")"
        ")"
    Alias ="LateBuild"
    Expression ="CDate(DMax(\"UnitBuildDate\",\"Units\",\"GroupREF=\" & [SchGrp].[SGrpID] & \"\")"
        ")"
    Alias ="EarlyActualDelivery"
    Expression ="CDate(DMin(\"UnitActualDelDate\",\"Units\",\"GroupREF=\" & [SchGrp].[SGrpID] & \""
        "\"))"
    Alias ="LateActualDelivery"
    Expression ="CDate(DMax(\"UnitActualDelDate\",\"Units\",\"GroupREF=\" & [SchGrp].[SGrpID] & \""
        "\"))"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Vendors"
    Expression ="SchGrp.POVendor = Vendors.VendorID"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Clients"
    Expression ="MstrLease.Client = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchFAID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Schedule"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="POnumber"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="POSent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EarlyEstimated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EarlyBuild"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LateBuild"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EarlyActualDelivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LateActualDelivery"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =118
    Top =21
    Right =1375
    Bottom =735
    Left =-1
    Top =-1
    Right =1225
    Bottom =442
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =26
        Top =40
        Right =170
        Bottom =278
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =199
        Top =68
        Right =386
        Bottom =456
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =415
        Top =196
        Right =559
        Bottom =340
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =587
        Top =318
        Right =733
        Bottom =457
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =586
        Top =56
        Right =730
        Bottom =200
        Top =0
        Name ="Clients"
        Name =""
    End
End

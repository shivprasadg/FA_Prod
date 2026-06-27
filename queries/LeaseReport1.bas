Operation =1
Option =0
Where ="(((MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Alias ="Lease No"
    Expression ="MstrLease.MLNo"
    Alias ="MstrLeaseDate"
    Expression ="MstrLease.MLDt"
    Expression ="MstrLease.MLOrig"
    Alias ="FAID"
    Expression ="Schedule.SchFAID"
    Alias ="POSentDate"
    Expression ="Schedule.POSentDt"
    Alias ="AssetId"
    Expression ="Units.UnitID"
    Alias ="ModelYear"
    Expression ="Units.UnitYr"
    Alias ="Make"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitType"
    Expression ="Schedule.SchNo"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="Unit No"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.Unitvin"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType] & IIf([Axle]=\"tbd\",Null,\" \" & [axle])"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="Co_Lessee"
    Expression ="Clients.clientCoNm2"
    Alias ="Title Num"
    Expression ="Units.UnitTitleNumber"
    Alias ="Status"
    Expression ="Units.unitstatus"
    Alias ="InService"
    Expression ="Units.UnitInServ"
    Alias ="UnitCost"
    Expression ="CCur(Nz([UnitFinalCost],0))"
    Expression ="Units.UnitInvoiceNum"
    Alias ="UnitInvoicePaid"
    Expression ="Units.UnitVendorPytDate"
    Expression ="Units.UnitCOABack"
    Expression ="Units.UnitEstDelDate"
    Alias ="SynDate"
    Expression ="SchGrp.FundDate"
    Alias ="PODate"
    Expression ="SchGrp.POSentDt"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.POSentDt"
    Flag =1
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
        dbText "Name" ="Units.Unitvin"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Unit No"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1485"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="3300"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Co_Lessee"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" =">"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="InService"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Title Num"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLeaseDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="POSentDate"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SynDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PODate"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-201
    Top =38
    Right =1145
    Bottom =947
    Left =-1
    Top =-1
    Right =1314
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =33
        Top =16
        Right =175
        Bottom =200
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =199
        Top =16
        Right =340
        Bottom =160
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =377
        Top =18
        Right =521
        Bottom =200
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =551
        Top =16
        Right =695
        Bottom =160
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =712
        Top =12
        Right =893
        Bottom =380
        Top =0
        Name ="Units"
        Name =""
    End
End

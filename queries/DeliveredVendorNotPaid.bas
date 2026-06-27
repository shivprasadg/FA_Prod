Operation =1
Option =0
Where ="(((vw_CustDelvReportBase.FundDate) Is Null) AND (Not (vw_CustDelvReportBase.Actu"
    "alDeliveryDate) Is Null) AND ((vw_CustDelvReportBase.VendorPaidDate) Is Null) AN"
    "D ((vw_CustDelvReportBase.BankShortName)<>\"FA LLC\"))"
Begin InputTables
    Name ="vw_CustDelvReportBase"
End
Begin OutputColumns
    Expression ="vw_CustDelvReportBase.ClientShortName"
    Expression ="vw_CustDelvReportBase.Schedule"
    Expression ="vw_CustDelvReportBase.UnitGroup"
    Expression ="vw_CustDelvReportBase.AssetID"
    Expression ="vw_CustDelvReportBase.UnitDescAbbr"
    Expression ="vw_CustDelvReportBase.VendorSNm"
    Expression ="vw_CustDelvReportBase.FundDate"
    Expression ="vw_CustDelvReportBase.ActualDeliveryDate"
    Expression ="vw_CustDelvReportBase.[Interim Bank]"
    Expression ="vw_CustDelvReportBase.VendorPaidDate"
    Alias ="DaysDelivered"
    Expression ="DateDiff(\"d\",[ActualDeliveryDate],Date())"
    Expression ="vw_CustDelvReportBase.BankShortName"
End
Begin OrderBy
    Expression ="vw_CustDelvReportBase.ClientShortName"
    Flag =0
    Expression ="vw_CustDelvReportBase.Schedule"
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
        dbText "Name" ="vw_CustDelvReportBase.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.UnitDescAbbr"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.VendorSNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.[Interim Bank]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.VendorPaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustDelvReportBase.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysDelivered"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =25
    Top =169
    Right =1410
    Bottom =1201
    Left =-1
    Top =-1
    Right =1361
    Bottom =452
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =278
        Top =154
        Right =752
        Bottom =707
        Top =0
        Name ="vw_CustDelvReportBase"
        Name =""
    End
End

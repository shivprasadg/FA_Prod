Operation =1
Option =0
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Alias ="Account Name"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Alias ="Lease Sch"
    Expression ="RemarketingReportAllUnits.Schedule"
    Alias ="Group"
    Expression ="RemarketingReportAllUnits.Unitgroup"
    Alias ="AssetID"
    Expression ="RemarketingReportAllUnits.UnitID"
    Alias ="Customer Unit#"
    Expression ="RemarketingReportAllUnits.UnitNum"
    Alias ="VIN"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Alias ="InvStatus"
    Expression ="RemarketingReportAllUnits.RStatus"
    Alias ="Year"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Alias ="Make"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Alias ="Model"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Alias ="Type"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.ParkDate"
    Expression ="RemarketingReportAllUnits.[SUF Sent]"
    Expression ="RemarketingReportAllUnits.[SUF Returned]"
    Alias ="DaysIn"
    Expression ="RemarketingReportAllUnits.DaysInInventory"
    Alias ="InspnDate"
    Expression ="RemarketingReportAllUnits.InspectedDate"
    Alias ="InspPassed"
    Expression ="RemarketingReportAllUnits.InspectionPassed"
    Expression ="RemarketingReportAllUnits.OffLeaseDate"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Alias ="PayOff Sent"
    Expression ="RemarketingReportAllUnits.PayOffRequested"
    Alias ="PayOff Back"
    Expression ="RemarketingReportAllUnits.PayOffReceived"
    Alias ="Title Recv'd"
    Expression ="RemarketingReportAllUnits.TitleRecvHist"
    Expression ="RemarketingReportAllUnits.Buyer"
    Expression ="RemarketingReportAllUnits.SalesPerson"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
    Alias ="Buy Out Cost"
    Expression ="RemarketingReportAllUnits.BuyOut"
    Expression ="RemarketingReportAllUnits.UnitSalePrice"
    Expression ="RemarketingReportAllUnits.TermPenalty"
    Alias ="RMKTExpense"
    Expression ="RemarketingReportAllUnits.TotalExpense"
    Expression ="RemarketingReportAllUnits.UnitProfit"
    Expression ="RemarketingReportAllUnits.BankPartner"
    Expression ="RemarketingReportAllUnits.ParkedLocation"
    Alias ="LAddr"
    Expression ="[ParkedAddress] & \" \" & [PCity] & \", \" & [PState]"
    Expression ="RemarketingReportAllUnits.PCity"
    Expression ="RemarketingReportAllUnits.PState"
    Expression ="RemarketingReportAllUnits.ClientID"
    Expression ="RemarketingReportAllUnits.BuyerID"
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.LocationID"
    Expression ="RemarketingReportAllUnits.AH_Vin"
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
        dbText "Name" ="RemarketingReportAllUnits.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Account Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Daysin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BankPartner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PayOff Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease Sch"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PayOFf Back"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buy Out Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Recv'd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SalesPerson"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitSalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspPassed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTExpense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Laddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspnDate"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitProfit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.AH_Vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.pState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.PCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkedLocation"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =172
    Top =35
    Right =1581
    Bottom =1266
    Left =-1
    Top =-1
    Right =1385
    Bottom =206
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =119
        Top =3
        Right =412
        Bottom =228
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End

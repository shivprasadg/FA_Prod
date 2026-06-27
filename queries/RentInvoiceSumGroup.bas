Operation =1
Option =0
Begin InputTables
    Name ="vw_RentInvoice"
End
Begin OutputColumns
    Expression ="vw_RentInvoice.BillID"
    Expression ="vw_RentInvoice.GroupID"
    Expression ="vw_RentInvoice.Schedule"
    Expression ="vw_RentInvoice.UnitGroup"
    Expression ="vw_RentInvoice.LocationName"
    Expression ="vw_RentInvoice.LCity"
    Expression ="vw_RentInvoice.LState"
    Expression ="vw_RentInvoice.StTaxType"
    Alias ="UnitCount"
    Expression ="Sum(vw_RentInvoice.CountMe)"
    Alias ="BilledRent"
    Expression ="Sum(vw_RentInvoice.BilledRent)"
    Alias ="BilledFees"
    Expression ="Sum(vw_RentInvoice.BilledFees)"
    Alias ="UnitTaxes"
    Expression ="Sum(vw_RentInvoice.UnitTaxes)"
    Alias ="UnitTotalRent"
    Expression ="Sum(vw_RentInvoice.UnitTotalRent)"
End
Begin Groups
    Expression ="vw_RentInvoice.BillID"
    GroupLevel =0
    Expression ="vw_RentInvoice.GroupID"
    GroupLevel =0
    Expression ="vw_RentInvoice.Schedule"
    GroupLevel =0
    Expression ="vw_RentInvoice.UnitGroup"
    GroupLevel =0
    Expression ="vw_RentInvoice.LocationName"
    GroupLevel =0
    Expression ="vw_RentInvoice.LCity"
    GroupLevel =0
    Expression ="vw_RentInvoice.LState"
    GroupLevel =0
    Expression ="vw_RentInvoice.StTaxType"
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
        dbText "Name" ="BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.UnitTaxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledFees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.BilledFees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.UnitTotalRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.CountMe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCountMe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.BilledRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitTotalRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfBilledFees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitTaxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTaxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfBilledRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.[sttaxtype]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.sttaxtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfUnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RentInvoice.GroupID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =74
    Top =444
    Right =1179
    Bottom =1071
    Left =-1
    Top =-1
    Right =1081
    Bottom =262
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =300
        Top =0
        Name ="vw_RentInvoice"
        Name =""
    End
End

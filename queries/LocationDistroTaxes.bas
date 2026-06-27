Operation =1
Option =0
Begin InputTables
    Name ="vw_LocationDistribution_Taxes"
End
Begin OutputColumns
    Expression ="vw_LocationDistribution_Taxes.Client"
    Expression ="vw_LocationDistribution_Taxes.LocationName"
    Expression ="vw_LocationDistribution_Taxes.Locationsn"
    Expression ="vw_LocationDistribution_Taxes.LCity"
    Expression ="vw_LocationDistribution_Taxes.LState"
    Expression ="vw_LocationDistribution_Taxes.LCounty"
    Expression ="vw_LocationDistribution_Taxes.LZip"
    Alias ="LUnits"
    Expression ="Format([vw_LocationDistribution_Taxes].[LUnits],\"#,##0\")"
    Alias ="FAUnits"
    Expression ="Format([vw_LocationDistribution_Taxes].[FAUnits],\"#,##0\")"
    Alias ="Distro"
    Expression ="Format([vw_LocationDistribution_Taxes].[Distro],\"Percent\")"
    Alias ="LTaxRate"
    Expression ="Format([vw_LocationDistribution_Taxes].[LTaxRate],\"Percent\")"
    Expression ="vw_LocationDistribution_Taxes.LTaxOption"
    Alias ="LocTaxes"
    Expression ="Format([vw_LocationDistribution_Taxes].[LocTaxes],\"Percent\")"
    Expression ="vw_LocationDistribution_Taxes.StateSales"
    Expression ="vw_LocationDistribution_Taxes.CountySales"
    Expression ="vw_LocationDistribution_Taxes.CitySales"
    Expression ="vw_LocationDistribution_Taxes.StateRReceipt"
    Expression ="vw_LocationDistribution_Taxes.LStateID"
    Expression ="vw_LocationDistribution_Taxes.LCountyID"
    Expression ="vw_LocationDistribution_Taxes.LCityID"
    Expression ="vw_LocationDistribution_Taxes.LocationID"
    Expression ="vw_LocationDistribution_Taxes.ClientGroupId"
    Expression ="vw_LocationDistribution_Taxes.clientID"
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
        dbText "Name" ="vw_LocationDistribution_Taxes.LTaxOption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.StateSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Distro"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.StateRReceipt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.CountySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.CitySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LCityID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocTaxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationDistribution_Taxes.Locationsn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =156
    Top =84
    Right =2487
    Bottom =1291
    Left =-1
    Top =-1
    Right =2307
    Bottom =842
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_LocationDistribution_Taxes"
        Name =""
    End
End

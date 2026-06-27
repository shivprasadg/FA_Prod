Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="vw_LocationTaxes"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="CountOfAssetID"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="PaysTax"
    Expression ="IIf([UnitTaxExempt]=False,Count([AssetID]),0)"
    Alias ="TaxExempt"
    Expression ="IIf([UnitTaxExempt]=True,Count([AssetID]),0)"
    Expression ="vw_SixKeys.LCity"
    Expression ="vw_SixKeys.LState"
    Alias ="SalesTaxRate"
    Expression ="[StateSales]+[CountySales]+[CitySales]"
    Expression ="Units.[UnitTaxExempt]"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="vw_LocationTaxes"
    Expression ="vw_SixKeys.LocationID = vw_LocationTaxes.LocationID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.LCity"
    GroupLevel =0
    Expression ="vw_SixKeys.LState"
    GroupLevel =0
    Expression ="[StateSales]+[CountySales]+[CitySales]"
    GroupLevel =0
    Expression ="Units.[UnitTaxExempt]"
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
        dbText "Name" ="vw_LocationTaxes.CitySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PaysTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationTaxes.StateSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_LocationTaxes.CountySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.LocationState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitTaxExempt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxExempt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =111
    Top =73
    Right =1632
    Bottom =1059
    Left =-1
    Top =-1
    Right =1503
    Bottom =709
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =449
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =420
        Top =142
        Right =655
        Bottom =593
        Top =0
        Name ="vw_LocationTaxes"
        Name =""
    End
    Begin
        Left =229
        Top =75
        Right =373
        Bottom =219
        Top =0
        Name ="Units"
        Name =""
    End
End

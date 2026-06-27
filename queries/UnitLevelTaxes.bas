Operation =1
Option =0
Where ="(((Units.UnitStatus)=\"A\" Or (Units.UnitStatus)=\"B\") AND ((TaxCitiesBase.TaxT"
    "ypeID)=1 Or (TaxCitiesBase.TaxTypeID) Is Null) AND ((TaxCountiesBase.TaxTypeId)="
    "1 Or (TaxCountiesBase.TaxTypeId) Is Null) AND ((TaxStatesBase.TaxTypeId)=1 Or (T"
    "axStatesBase.TaxTypeId) Is Null))"
Begin InputTables
    Name ="Units"
    Name ="Locations"
    Name ="SchGrp"
    Name ="TaxCitiesBase"
    Name ="TaxCountiesBase"
    Name ="TaxStatesBase"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitStatus"
    Alias ="InServiceDate"
    Expression ="Units.UnitInServ"
    Alias ="PDStartDate"
    Expression ="Units.UnitPDStartDate"
    Alias ="Rent"
    Expression ="CCur(IIf([UnitRent]>0,[UnitRent],[UnitSRent]))"
    Alias ="FeeNoTax"
    Expression ="CCur(Nz([SGrpFSLPmt],0))"
    Alias ="FeeTaxable"
    Expression ="CCur(Nz([SGrpRSLTPmt],0))"
    Alias ="UnitTaxes"
    Expression ="CCur(Round(IIf([UnitTaxExempt]=True,0,([UnitRent]+[FeeTaxable])*[LTaxRate]),2))"
    Alias ="RentPayment"
    Expression ="[Rent]+[FeeNoTax]+[FeeTaxable]+[UnitTaxes]"
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LTaxRate"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Alias ="StateRate"
    Expression ="TaxStatesBase.BaseTaxRate"
    Alias ="CountyRate"
    Expression ="TaxCountiesBase.BaseTaxRate"
    Alias ="CityRate"
    Expression ="TaxCitiesBase.BaseTaxRate"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="TaxCitiesBase"
    Expression ="Locations.LCityID = TaxCitiesBase.TaxCityID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="TaxCountiesBase"
    Expression ="Locations.LCountyID = TaxCountiesBase.TaxCountyID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="TaxStatesBase"
    Expression ="Locations.LStateID = TaxStatesBase.TaxStatesID"
    Flag =2
End
Begin OrderBy
    Expression ="Locations.Locationsn"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="UnitTaxes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="InServiceDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FeeNoTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="FeeTaxable"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="RentPayment"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="PDStartDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="StateRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="CountyRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CityRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
End
Begin
    State =0
    Left =-125
    Top =-3
    Right =1265
    Bottom =922
    Left =-1
    Top =-1
    Right =1372
    Bottom =440
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =250
        Bottom =562
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =601
        Top =40
        Right =786
        Bottom =334
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =357
        Top =37
        Right =541
        Bottom =287
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =882
        Top =192
        Right =1026
        Bottom =437
        Top =0
        Name ="TaxCitiesBase"
        Name =""
    End
    Begin
        Left =1010
        Top =76
        Right =1154
        Bottom =220
        Top =0
        Name ="TaxCountiesBase"
        Name =""
    End
    Begin
        Left =823
        Top =6
        Right =967
        Bottom =150
        Top =0
        Name ="TaxStatesBase"
        Name =""
    End
End

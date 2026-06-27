Operation =1
Option =0
Begin InputTables
    Name ="TaxStates"
    Name ="TaxCounties"
    Name ="TaxCities"
End
Begin OutputColumns
    Expression ="TaxStates.ISOCountryCode2"
    Alias ="TaxStateID"
    Expression ="TaxStates.ID"
    Alias ="TaxCountyID"
    Expression ="TaxCounties.ID"
    Alias ="TaxCityID"
    Expression ="TaxCities.ID"
    Alias ="State"
    Expression ="TaxStates.PostalAbbreviation"
    Alias ="County"
    Expression ="TaxCounties.CountyName"
    Alias ="City"
    Expression ="TaxCities.CityName"
End
Begin Joins
    LeftTable ="TaxStates"
    RightTable ="TaxCounties"
    Expression ="TaxStates.ID = TaxCounties.TaxStatesID"
    Flag =1
    LeftTable ="TaxCounties"
    RightTable ="TaxCities"
    Expression ="TaxCounties.ID = TaxCities.TaxCountyID"
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
        dbText "Name" ="TaxCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxCityID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="County"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxStates.ISOCountryCode2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =0
    Right =1249
    Bottom =969
    Left =-1
    Top =-1
    Right =1153
    Bottom =675
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =56
        Top =66
        Right =200
        Bottom =210
        Top =0
        Name ="TaxStates"
        Name =""
    End
    Begin
        Left =248
        Top =125
        Right =392
        Bottom =269
        Top =0
        Name ="TaxCounties"
        Name =""
    End
    Begin
        Left =437
        Top =217
        Right =581
        Bottom =361
        Top =0
        Name ="TaxCities"
        Name =""
    End
End

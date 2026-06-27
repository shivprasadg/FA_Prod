Operation =2
Name ="LocationsCountyTemp"
Option =0
Where ="(((Locations.LState)=[AskState]) AND ((Locations.LocationType)<>\"Title\"))"
Begin InputTables
    Name ="TaxCounties"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.LocationID"
    Expression ="Locations.LState"
    Expression ="Locations.LCity"
    Expression ="Locations.LZip"
    Expression ="TaxCounties.CountyName"
    Expression ="Locations.LCounty"
    Expression ="Locations.LCountyID"
    Expression ="Locations.LStateID"
    Expression ="Locations.LocationType"
    Alias ="LookCounty"
    Expression ="DLookUp(\"County\",\"TaxZipFipsMatrix\",\"Zip='\" & Left([LZip],5) & \"'\")"
    Alias ="LookCountyID"
    Expression ="DLookUp(\"ID\",\"TaxCounties\",\"TaxStatesID=\" & [AskStateId] & \"  AND CountyN"
        "ame='\" & [LookCounty] & \"'\")"
End
Begin Joins
    LeftTable ="TaxCounties"
    RightTable ="Locations"
    Expression ="TaxCounties.TaxStatesID = Locations.LStateID"
    Flag =1
    LeftTable ="TaxCounties"
    RightTable ="Locations"
    Expression ="TaxCounties.CountyName = Locations.LCounty"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Locations.LStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxCounties.TaxStatesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxCounties.CountyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookCountyID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-94
    Top =64
    Right =1428
    Bottom =834
    Left =-1
    Top =-1
    Right =1504
    Bottom =425
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =21
        Right =204
        Bottom =344
        Top =0
        Name ="TaxCounties"
        Name =""
    End
    Begin
        Left =350
        Top =-10
        Right =614
        Bottom =382
        Top =0
        Name ="Locations"
        Name =""
    End
End

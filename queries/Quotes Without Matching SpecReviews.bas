Operation =1
Option =0
Where ="(((SpecReviews.SRID) Is Null))"
Begin InputTables
    Name ="Quotes"
    Name ="SpecReviews"
End
Begin OutputColumns
    Expression ="SpecReviews.SRID"
    Expression ="Quotes.specREF"
End
Begin Joins
    LeftTable ="Quotes"
    RightTable ="SpecReviews"
    Expression ="Quotes.specREF = SpecReviews.SRID"
    Flag =2
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
        dbText "Name" ="[Quotes].[specREF]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.specREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpecReviews.SRID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1177
    Bottom =702
    Left =-1
    Top =-1
    Right =1145
    Bottom =223
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =218
        Bottom =227
        Top =0
        Name ="Quotes"
        Name =""
    End
    Begin
        Left =263
        Top =13
        Right =452
        Bottom =234
        Top =0
        Name ="SpecReviews"
        Name =""
    End
End

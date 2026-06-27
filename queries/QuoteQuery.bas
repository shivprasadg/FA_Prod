Operation =1
Option =0
Begin InputTables
    Name ="Quotes"
End
Begin OutputColumns
    Expression ="Quotes.QuoteID"
    Expression ="Quotes.groupREF"
    Expression ="Quotes.dealRef"
    Expression ="Quotes.QuoteDate"
    Expression ="Quotes.QuoteBLCD"
    Expression ="Quotes.QuoteLeaseTerm"
    Expression ="Quotes.QuoteLRF"
    Expression ="Quotes.QuotePayment"
    Expression ="Quotes.QuoteBonusDep"
    Expression ="Quotes.QuoteIndex"
    Expression ="Quotes.QuoteIndexYear"
    Expression ="Quotes.QuoteIndexDate"
    Expression ="Quotes.QuoteIndexRate"
    Expression ="Quotes.QuoteMPYHPY"
    Expression ="Quotes.QuoteOverMiCharge"
    Expression ="Quotes.QuoteAdvArr"
    Expression ="Quotes.QuoteYears"
    Expression ="Quotes.QuoteTotal"
    Expression ="Quotes.QuoteGrandTotal"
    Expression ="Quotes.QuoteOptMo1"
    Expression ="Quotes.QuoteOptMo2"
    Expression ="Quotes.QuoteOptMo3"
    Expression ="Quotes.QuoteOption1"
    Expression ="Quotes.QuoteOption2"
    Expression ="Quotes.QuoteOption3"
    Expression ="Quotes.QuoteOptRate1"
    Expression ="Quotes.QuoteOptRate2"
    Expression ="Quotes.QuoteOptRate3"
    Expression ="Quotes.QuoteRequote"
    Expression ="Quotes.QuoteNotes"
    Expression ="Quotes.QuoteFinal"
    Expression ="Quotes.QuoteOpt1Exercised"
    Expression ="Quotes.QuoteOpt2Exercised"
    Expression ="Quotes.QuoteOpt3Exercised"
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
        dbText "Name" ="Quotes.QuoteOptRate3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteLeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteFinal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteOpt1Exercised"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteOpt3Exercised"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quotes.QuoteOpt2Exercised"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =404
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =407
        Bottom =334
        Top =0
        Name ="Quotes"
        Name =""
    End
End

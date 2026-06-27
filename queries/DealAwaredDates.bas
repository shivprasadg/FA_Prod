Operation =1
Option =0
Having ="(((Deals.dealLPAFSent) Is Not Null) AND ((Deals.dealOutcome)=\"awarded\"))"
Begin InputTables
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Deals.FAID"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealOutcome"
    Alias ="DealYear"
    Expression ="DatePart(\"yyyy\",[dealOutcomeDate])"
End
Begin OrderBy
    Expression ="Deals.dealLPAFSent"
    Flag =1
End
Begin Groups
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Deals.dealLPAFSent"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[dealOutcomeDate])"
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
        dbText "Name" ="Deals.dealLPAFSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-466
    Top =117
    Right =1049
    Bottom =857
    Left =-1
    Top =-1
    Right =1186
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =63
        Top =14
        Right =444
        Bottom =368
        Top =0
        Name ="Deals"
        Name =""
    End
End

Operation =4
Option =0
Begin InputTables
    Name ="TitleLogTemp"
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.UnitTitleNumber"
    Expression ="[TitleLogTemp].[Titlenumber]"
    Name ="Units.UnitTitleRecd"
    Expression ="[TitleLogTemp].[DateLog]"
End
Begin Joins
    LeftTable ="TitleLogTemp"
    RightTable ="Units"
    Expression ="TitleLogTemp.UnitID = Units.UnitID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleLogTemp.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleLogTemp.DateLog"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1604
    Bottom =872
    Left =-1
    Top =-1
    Right =1572
    Bottom =532
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =264
        Top =154
        Right =408
        Bottom =298
        Top =0
        Name ="TitleLogTemp"
        Name =""
    End
    Begin
        Left =59
        Top =111
        Right =203
        Bottom =478
        Top =0
        Name ="Units"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="Deals"
    Name ="client_lookup"
    Name ="Groups"
End
Begin OutputColumns
    Expression ="Deals.FAID"
    Expression ="Groups.groupNo"
    Expression ="Groups.groupID"
    Expression ="Groups.dealRef"
    Expression ="client_lookup.CompanyAndDBA"
    Expression ="Groups.groupUnits"
    Expression ="Groups.groupYear"
    Expression ="Groups.groupMakeRef"
    Expression ="Groups.groupModelRef"
    Expression ="Groups.groupTypeRef"
    Expression ="Groups.groupTotal"
    Expression ="Groups.groupOptMo1"
    Expression ="Groups.groupOptMo2"
    Expression ="Groups.groupOptMo3"
    Expression ="Groups.groupOption1"
    Expression ="Groups.groupOption2"
    Expression ="Groups.groupOption3"
    Expression ="Groups.groupMPYHPY"
    Expression ="Deals.dealSchedNum"
    Expression ="Deals.dealShortDesc"
    Expression ="Deals.clientRef"
    Expression ="Groups.groupBillingRef"
    Expression ="Groups.groupBillingNotes"
    Expression ="Groups.GRPLOCSN"
    Expression ="Groups.AltEquip"
    Expression ="Groups.groupMake"
    Expression ="Groups.groupModel"
    Expression ="Groups.groupType"
    Expression ="Deals.DealPLB"
    Expression ="Groups.GrpSchedule"
    Expression ="Groups.groupTotal"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="client_lookup"
    Expression ="Deals.clientRef = client_lookup.clientID"
    Flag =1
    LeftTable ="Deals"
    RightTable ="Groups"
    Expression ="Deals.dealID = Groups.dealRef"
    Flag =1
End
Begin OrderBy
    Expression ="Deals.FAID"
    Flag =0
    Expression ="Groups.groupNo"
    Flag =0
    Expression ="Groups.groupID"
    Flag =0
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
        dbText "Name" ="Groups.groupMPYHPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="client_lookup.CompanyAndDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMakeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModelRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOptMo1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOptMo2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOptMo3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOption1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOption2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupOption3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.dealRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupBillingRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Groups.groupBillingNotes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Groups.GRPLOCSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.AltEquip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.DealPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.GrpSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1010"
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
    Bottom =286
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =387
        Top =44
        Right =531
        Bottom =188
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =620
        Top =20
        Right =764
        Bottom =164
        Top =0
        Name ="client_lookup"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =318
        Bottom =261
        Top =0
        Name ="Groups"
        Name =""
    End
End

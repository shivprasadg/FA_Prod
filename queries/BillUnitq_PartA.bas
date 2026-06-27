Operation =1
Option =0
Where ="(((Units.UnitStatus)<>\"b\") AND ((Units.UnitDocAccpt)=False)) OR (((Units.UnitI"
    "nServInd)=\"a\") AND ((Units.UnitStatus)<>\"b\") AND ((Units.UnitDocAccpt)=True)"
    " AND ((Units.UnitInServ) Is Not Null)) OR (((Units.UnitInServInd)=\"a\") AND ((U"
    "nits.UnitStatus)<>\"b\") AND ((Units.UnitDocAccpt)=True) AND ((Units.UnitInServ)"
    " Is Not Null) AND ((Units.UnitPDStartDate) Is Not Null)) OR (((Units.UnitInServI"
    "nd)=\"a\" And (Units.UnitInServInd) Is Not Null) AND ((Units.UnitStatus)=\"b\"))"
    " OR (((Units.UnitInServInd)=\"p\") AND ((Units.UnitStatus)=\"A\") AND ((Units.Un"
    "itPDStartDate) Is Not Null))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Bills"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpID"
    Expression ="Units.UnitID"
    Alias ="UBStart"
    Expression ="IIf(IsNull([unitpdstartdate]),IIf(IsNull([unitinserv]),[unitacceptdate],[unitins"
        "erv]),[unitpdstartdate])"
    Alias ="unitmax"
    Expression ="SchGrp.BLCD"
    Expression ="Units.UnitInServInd"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitDocAccpt"
    Alias ="UnitPDST"
    Expression ="IIf([sgrpnoslv]=No,IIf([UBSTART]>[FORMS]![BillingSubForm]![BILLTO],[FORMS]![Bill"
        "ingSubForm]![BILLTO]+1,IIf([UBSTART]>[FORMS]![BillingSubForm]![BILLFROM],[UBSTAR"
        "T],[FORMS]![BillingSubForm]![BILLFROM])),IIf([ubstart]>=[schcrexp]+1 And [ubstar"
        "t]>=[forms]![BillingSubForm]![billfrom],[ubstart],IIf([schcrexp]+1>=[forms]![Bil"
        "lingSubForm]![billfrom],[schcrexp]+1,[forms]![BillingSubForm]![billfrom])))"
    Expression ="Units.UnitInServ"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="Bills"
    Expression ="Schedule.SchID = Bills.BillSchRef"
    Flag =1
    LeftTable ="Bills"
    RightTable ="SchGrp"
    Expression ="Bills.BillSchRef = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpID"
    Flag =0
    Expression ="Units.UnitID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UBStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="unitmax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitPDST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =170
    Top =37
    Right =1428
    Bottom =683
    Left =-1
    Top =-1
    Right =1226
    Bottom =103
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =2
        Right =206
        Bottom =146
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =243
        Top =0
        Right =387
        Bottom =144
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =416
        Top =9
        Right =560
        Bottom =153
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =636
        Top =3
        Right =874
        Bottom =136
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =883
        Top =5
        Right =1101
        Bottom =185
        Top =0
        Name ="Units"
        Name =""
    End
End

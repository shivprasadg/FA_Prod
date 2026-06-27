Operation =1
Option =0
Where ="(((Bills.BillID)=[Forms]![Billing]![BillingSubForm].[Form]![BillID]))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Bills"
    Name ="SchGrp"
    Name ="Units"
    Name ="PerDiemBill_MostRecentPD2"
End
Begin OutputColumns
    Expression ="MstrLease.Client"
    Expression ="Bills.BillID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpID"
    Expression ="Units.UnitID"
    Alias ="UBStart"
    Expression ="IIf(IsNull([unitpdstartdate]) And IsNull([unitinserv]),[unitacceptdate],Nz([unit"
        "pdstartdate],[unitinserv]))"
    Alias ="unitmax"
    Expression ="SchGrp.BLCD"
    Expression ="Units.UnitInServInd"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitDocAccpt"
    Alias ="UnitPDSTOld"
    Expression ="IIf([sgrpnoslv]=No,IIf([UBSTART]>[BILLTO],[BILLTO]+1,IIf([UBSTART]>[BILLFROM],[U"
        "BSTART],[BILLFROM])),IIf([ubstart]>=[schcrexp]+1 And [ubstart]>=[billfrom],[ubst"
        "art],IIf([schcrexp]+1>=[billfrom],[schcrexp]+1,[billfrom])))"
    Alias ="UnitPDST"
    Expression ="IIf([UBStart]>=[PSent] And [UBStart]<=[PStart] And [Client]=4971,[UBStart],IIf(["
        "sgrpnoslv]=No,IIf([UBSTART]>[BILLTO],[BILLTO]+1,IIf([UBSTART]>[BILLFROM],[UBSTAR"
        "T],[BILLFROM])),IIf([ubstart]>=[schcrexp]+1 And [ubstart]>=[billfrom],[ubstart],"
        "IIf([schcrexp]+1>=[billfrom],[schcrexp]+1,[billfrom]))))"
    Expression ="Units.UnitInServ"
    Expression ="Bills.Billfrom"
    Expression ="Schedule.SchID"
    Alias ="PBillDate"
    Expression ="CDate(Nz([BillDx],#1/1/1900#))"
    Alias ="PSent"
    Expression ="CDate(Nz([Sent],#1/1/1910#))"
    Alias ="PStart"
    Expression ="CDate(Nz([Start],#1/1/4500#))"
    Alias ="PEnd"
    Expression ="CDate(Nz([End],#12/31/4500#))"
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
    LeftTable ="Schedule"
    RightTable ="PerDiemBill_MostRecentPD2"
    Expression ="Schedule.SchID = PerDiemBill_MostRecentPD2.SchID"
    Flag =2
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
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="UBStart"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="unitmax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="UnitPDST"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Bills.Billfrom"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="PStart"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="PSent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="PEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitPDSTOld"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PBillDate"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-467
    Top =186
    Right =1484
    Bottom =832
    Left =-1
    Top =-1
    Right =1919
    Bottom =54
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
        Left =596
        Top =20
        Right =757
        Bottom =223
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =796
        Top =54
        Right =964
        Bottom =234
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1023
        Top =13
        Right =1167
        Bottom =157
        Top =0
        Name ="PerDiemBill_MostRecentPD2"
        Name =""
    End
End

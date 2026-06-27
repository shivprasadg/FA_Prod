Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="SchGrp"
    Name ="Schedule"
    Name ="Lookup_LeaseStartsOn"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Fleet_Advantage_Agreement"
    Expression ="\"Sch \" & [Schedule] & \" - Grp \" & [UnitGroup]"
    Alias ="Assignee Contract No"
    Expression ="Nz([SchBnkClNO],\"\") & IIf(Nz([SgrpBnkNo],\"\")<>\"\",\"-\" & [SgrpBnkNo],\"\")"
    Alias ="Assignee Bank"
    Expression ="vw_SixKeys.BankShortName"
    Alias ="Client Unit Num"
    Expression ="vw_SixKeys.UnitNum"
    Expression ="vw_SixKeys.VIN"
    Alias ="UnitDesc"
    Expression ="vw_SixKeys.UnitDescDocVer"
    Alias ="Location"
    Expression ="[LCity] & \" \" & [LState]"
    Alias ="Unit Cost"
    Expression ="CCur([UnitCost])"
    Alias ="Child Cost"
    Expression ="CCur([PartsCost])"
    Alias ="Total Cost"
    Expression ="[UnitCost]+[PartsCost]"
    Alias ="Commencement_Date"
    Expression ="vw_SixKeys.BLCD"
    Expression ="vw_SixKeys.AcceptedDate"
    Alias ="PerDiemStart"
    Expression ="Lookup_LeaseStartsOn.StartsOn"
    Alias ="PerDiemStartDate"
    Expression ="Units.UnitPDStartDate"
    Expression ="vw_SixKeys.InServiceDate"
    Expression ="vw_SixKeys.BLCD"
    Alias ="PDDays"
    Expression ="IIf(IsDate([UnitPDStartDate]) And IsDate([vw_SixKeys].[BLCD]),DateDiff(\"d\",[Un"
        "itPDStartDate],[vw_SixKeys].[BLCD]),0)"
    Alias ="PDRent"
    Expression ="Units.UNITPDRENT"
    Alias ="EstPerdiemTotal"
    Expression ="Nz([PDDays],0)*Nz([UnitPDRent],0)"
    Alias ="Final_Payment_Date"
    Expression ="vw_SixKeys.ProjEOL"
    Alias ="SchRent"
    Expression ="Units.unitsrent"
    Alias ="SyndRent"
    Expression ="Units.unitrent"
    Alias ="Lease_Rate Factor"
    Expression ="FormatPercent(IIf([UnitRent]>0,Round([UnitRent]/[Total Cost],7),0),4)"
    Alias ="Primary Lease Term"
    Expression ="[SchGrp].[LeaseTermPrimary]-Nz([SchGrp].[LeaseTermNotice],0)"
    Alias ="NoticeReq"
    Expression ="Nz([SchGrp].[LeaseTermNotice],0)"
    Alias ="MthsLeft"
    Expression ="IIf(IsDate([OffleaseActual]),0,DateDiff(\"m\",Date(),[ProjEOL]))"
    Alias ="Remain Rent"
    Expression ="IIf(IsDate([OffleaseActual]),0,DateDiff(\"m\",Date(),[ProjEOL])*[CurrRent])"
    Alias ="Delivered"
    Expression ="Units.UnitActualDelDate"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.Status"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Schedule"
    Expression ="vw_SixKeys.SchID = Schedule.SchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Lookup_LeaseStartsOn"
    Expression ="SchGrp.PerDiemStartsOn = Lookup_LeaseStartsOn.ID"
    Flag =2
End
Begin OrderBy
    Expression ="\"Sch \" & [Schedule] & \" - Grp \" & [UnitGroup]"
    Flag =0
    Expression ="[LCity] & \" \" & [LState]"
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
dbSingle "ECLScaleFactor" ="1"
Begin
    Begin
        dbText "Name" ="Unit Cost"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Primary Lease Term"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child Cost"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeReq"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MthsLeft"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Total Cost"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remain Rent"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Final_Payment_Date"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="29"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fleet_Advantage_Agreement"
        dbInteger "ColumnWidth" ="3150"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2190"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client Unit Num"
        dbInteger "ColumnWidth" ="1905"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lease_Rate Factor"
        dbInteger "ColumnWidth" ="2085"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Commencement_Date"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="2970"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee Contract No"
        dbInteger "ColumnWidth" ="2400"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee Bank"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5295"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PDDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="SchRent"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="PerDiemStart"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="vw_SixKeys.InServiceDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AcceptedDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PerDiemStartDate"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="EstPerdiemTotal"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.PartsCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PerDiemStartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lookup_LeaseStartsOn.StartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delivered"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1138
    Top =169
    Right =1052
    Bottom =1023
    Left =-1
    Top =-1
    Right =2166
    Bottom =244
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =366
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =242
        Top =30
        Right =448
        Bottom =283
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =704
        Top =-11
        Right =905
        Bottom =218
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =499
        Top =13
        Right =664
        Bottom =178
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =681
        Top =174
        Right =846
        Bottom =339
        Top =0
        Name ="Lookup_LeaseStartsOn"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((Locations.LocationID)=[Forms]![Schedule_Form]![Location]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="SchGrp"
    Name ="Locations"
    Name ="Schedule"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="lookup_States"
    Name ="Bank_Assignee"
    Name ="FSLFee_EndDate"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Assignee"
    Expression ="Bank_Assignee.ShortName"
    Expression ="Bank_Assignee.BankName"
    Expression ="vw_SixKeys.Schedule"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="UnitNum"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
    Alias ="Location"
    Expression ="Locations.LocationName"
    Alias ="Bank Client No"
    Expression ="[SchBnkClNO] & IIf(Left([SgrpBnkNo],1)=\"-\",[SgrpBnkNo],\"-\" & [SgrpBnkNo])"
    Alias ="Tax Exempt"
    Expression ="IIf([UnitTaxExempt]=-1,\"Exempt\",\"Taxed\")"
    Alias ="Tax Type"
    Expression ="lookup_States.StTaxType"
    Alias ="Tax Rate"
    Expression ="Locations.LTaxRate"
    Alias ="Rent"
    Expression ="CCur([unitsrent]+0)"
    Alias ="FSLFee"
    Expression ="CCur(Nz([SGrpFSLPmt],0))"
    Alias ="Tax Amt"
    Expression ="CCur(IIf([unittaxexempt]=-1,0,Round([locations].[LTaxRate]*[unitsrent],2)))"
    Alias ="Line Total"
    Expression ="[FslFee]+[Rent]+[Tax Amt]"
    Expression ="Locations.LocationID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Schedule"
    Expression ="vw_SixKeys.SchID = Schedule.SchID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="Locations_1"
    Expression ="Schedule.SchBillAdd = Locations_1.LocationID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Bank_Assignee"
    Expression ="SchGrp.Assignee_Bank = Bank_Assignee.ShortName"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="FSLFee_EndDate"
    Expression ="vw_SixKeys.SGrpID = FSLFee_EndDate.GroupId"
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
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3735"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank_Assignee.BankName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Bank Client No"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tax Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tax Exempt"
        dbInteger "ColumnWidth" ="1485"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tax Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tax Amt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="2370"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Line Total"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =53
    Top =128
    Right =1849
    Bottom =878
    Left =-1
    Top =-1
    Right =1764
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =18
        Top =12
        Right =162
        Bottom =462
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =187
        Top =231
        Right =345
        Bottom =637
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =368
        Top =172
        Right =513
        Bottom =363
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =368
        Top =397
        Right =512
        Bottom =652
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =367
        Top =17
        Right =511
        Bottom =161
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =542
        Top =184
        Right =686
        Bottom =328
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =538
        Top =413
        Right =682
        Bottom =557
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =862
        Top =279
        Right =1006
        Bottom =492
        Top =0
        Name ="Bank_Assignee"
        Name =""
    End
    Begin
        Left =861
        Top =44
        Right =1005
        Bottom =200
        Top =0
        Name ="FSLFee_EndDate"
        Name =""
    End
End

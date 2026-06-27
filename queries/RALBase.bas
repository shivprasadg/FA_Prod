Operation =1
Option =0
Having ="(((vw_SixKeys.SchID)=[Forms]![ScheduleControl]![SchID]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="SchGrp"
    Name ="Schedule"
    Name ="Banks"
    Name ="LocationsRemit"
    Alias ="LRemit"
    Name ="Clients"
    Name ="MstrLease"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.SchID"
    Alias ="RemitAddressID"
    Expression ="LRemit.locationId"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.MLNo"
    Alias ="MLDate"
    Expression ="Format([MLDt],\"mmmm d\"\", \"\"yyyy\")"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="LesseeAddress"
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & IIf(IsNull([clientCoNm2]),\"\",[client"
        "CoNm2] & Chr(13) & Chr(10)) & [clientAddress1] & \" \" & [clientAddress2] & Chr("
        "13) & Chr(10) & [clientcity] & \", \" & [clientstate] & \" \" & [clientzip]"
    Expression ="Banks.BankName"
    Alias ="BanksAddress"
    Expression ="[BankName] & Chr(13) & Chr(10) & [Address1] & Chr(13) & Chr(10) & IIf(Len([banks"
        "].[Address2])>3,[Address2] & Chr(13) & Chr(10),\"\") & [Banks].[City] & \", \" &"
        " [banks].[State] & \" \" & [banks].[Zip]"
    Alias ="BankAddressInLine"
    Expression ="[BankName] & \", \" & [Address1] & \", \" & IIf(Len([banks].[Address2])>3,\"Mail"
        " Stop \" & [Address2] & \", \",\"\") & [Banks].[City] & \", \" & [banks].[State]"
        " & \" \" & [banks].[Zip]"
    Alias ="BankRemitAddress"
    Expression ="[LRemit].[LocationStreet] & Chr(13) & Chr(10) & [LRemit].[LAddrLine1] & Chr(13) "
        "& Chr(10) & IIf(Len([LRemit].[LAddrLine2])>3,[LRemit].[LAddrLine2] & Chr(13) & C"
        "hr(10),\"\") & [LRemit].[LCity] & \" \" & [LRemit].[LState] & \" \" & Left([LRem"
        "it].[LZip],5) & IIf(Len([lremit].[lzip])>5,\"-\" & Right([lremit].[Lzip],4),\"\""
        ")"
    Alias ="GroupsIn"
    Expression ="GroupsNoinSchedule([vw_sixKeys].[schid])"
    Alias ="ConsistOf"
    Expression ="ScheduleGroupsInLine([vw_sixKeys].[schid])"
    Alias ="ScheduleDate"
    Expression ="Format([SchedDt],\"mmmm d\"\", \"\"yyyy\")"
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
    LeftTable ="vw_SixKeys"
    RightTable ="Banks"
    Expression ="vw_SixKeys.BankId = Banks.BankId"
    Flag =2
    LeftTable ="Banks"
    RightTable ="LRemit"
    Expression ="Banks.BankId = LRemit.ClientREF"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Clients"
    Expression ="vw_SixKeys.clientID = Clients.clientID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="MstrLease"
    Expression ="vw_SixKeys.MLID = MstrLease.MLID"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="LRemit.locationId"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="Format([MLDt],\"mmmm d\"\", \"\"yyyy\")"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & IIf(IsNull([clientCoNm2]),\"\",[client"
        "CoNm2] & Chr(13) & Chr(10)) & [clientAddress1] & \" \" & [clientAddress2] & Chr("
        "13) & Chr(10) & [clientcity] & \", \" & [clientstate] & \" \" & [clientzip]"
    GroupLevel =0
    Expression ="Banks.BankName"
    GroupLevel =0
    Expression ="[BankName] & Chr(13) & Chr(10) & [Address1] & Chr(13) & Chr(10) & IIf(Len([banks"
        "].[Address2])>3,[Address2] & Chr(13) & Chr(10),\"\") & [Banks].[City] & \", \" &"
        " [banks].[State] & \" \" & [banks].[Zip]"
    GroupLevel =0
    Expression ="[BankName] & \", \" & [Address1] & \", \" & IIf(Len([banks].[Address2])>3,\"Mail"
        " Stop \" & [Address2] & \", \",\"\") & [Banks].[City] & \", \" & [banks].[State]"
        " & \" \" & [banks].[Zip]"
    GroupLevel =0
    Expression ="[LRemit].[LocationStreet] & Chr(13) & Chr(10) & [LRemit].[LAddrLine1] & Chr(13) "
        "& Chr(10) & IIf(Len([LRemit].[LAddrLine2])>3,[LRemit].[LAddrLine2] & Chr(13) & C"
        "hr(10),\"\") & [LRemit].[LCity] & \" \" & [LRemit].[LState] & \" \" & Left([LRem"
        "it].[LZip],5) & IIf(Len([lremit].[lzip])>5,\"-\" & Right([lremit].[Lzip],4),\"\""
        ")"
    GroupLevel =0
    Expression ="Format([SchedDt],\"mmmm d\"\", \"\"yyyy\")"
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
dbInteger "RowHeight" ="1440"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BanksAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemitAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="3675"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConsistOf"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LesseeAddress"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLDate"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankAddressInLine"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankRemitAddress"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-239
    Top =55
    Right =1205
    Bottom =1049
    Left =-1
    Top =-1
    Right =1426
    Bottom =445
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =524
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="LRemit"
        Name =""
    End
    Begin
        Left =255
        Top =194
        Right =399
        Bottom =474
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =703
        Top =215
        Right =847
        Bottom =538
        Top =0
        Name ="MstrLease"
        Name =""
    End
End

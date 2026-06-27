Operation =1
Option =0
Where ="(((SchGrp.Assigned_Indicator)=\"F\") AND ((SchGrp.blcd) Is Not Null))"
Begin InputTables
    Name ="SchGrp"
    Name ="Schedule"
    Name ="Banks"
    Name ="LocationsRemit"
    Alias ="LRemit"
    Name ="Clients"
    Name ="MstrLease"
    Name ="BankPAI"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Alias ="RemitAddressID"
    Expression ="LRemit.locationId"
    Expression ="MstrLease.MLNo"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="GroupNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.UnitsInGroup"
    Alias ="MLDate"
    Expression ="Format([MLDt],\"mmmm d\"\", \"\"yyyy\")"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="LesseeAddress"
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & IIf(IsNull([clientCoNm2]),\"\",[client"
        "CoNm2] & Chr(13) & Chr(10)) & [clientAddress1] & \" \" & [clientAddress2] & Chr("
        "13) & Chr(10) & [clientcity] & \", \" & [clientstate] & \" \" & [clientzip]"
    Alias ="GroupRent"
    Expression ="SchGrp.LeasePymtForGroup"
    Alias ="RentDueDates"
    Expression ="DatePart(\"d\",[schgrp.BLCD]) & IIf(DatePart(\"d\",[schgrp.BLCD])=1,\"st\",IIf(D"
        "atePart(\"d\",[schgrp.BLCD])=31,\"st\",IIf(CStr(DatePart(\"d\",[schgrp.BLCD])) L"
        "ike '*2',\"nd\",IIf(DatePart(\"d\",[schgrp.BLCD])=3,\"rd\",\"th\")))) & \" of ea"
        "ch month [\" & [schgrp].[blcd] & \" ~ \" & IIf(IsNull([SchGrp.BLCD]),\"tbd\",IIf"
        "([LeasePaidAdvArr]=\"ADV\",DateAdd(\"m\",[schgrp].[LeaseTermPrimary]-1,[schgrp.B"
        "LCD]),DateAdd(\"m\",[schgrp].[LeaseTermPrimary],[schgrp.BLCD]))) & \"]\""
    Alias ="PayDate"
    Expression ="DatePart(\"d\",[schgrp.BLCD]) & IIf(DatePart(\"d\",[schgrp.BLCD])=1,\"st\",IIf(D"
        "atePart(\"d\",[schgrp.BLCD])=31,\"st\",IIf(CStr(DatePart(\"d\",[schgrp.BLCD])) L"
        "ike '*2',\"nd\",IIf(DatePart(\"d\",[schgrp.BLCD])=3,\"rd\",\"th\"))))"
    Expression ="Banks.BankShortName"
    Expression ="BankPAI.ABANumber"
    Expression ="BankPAI.AccountName"
    Expression ="BankPAI.AccountNumber"
    Expression ="BankPAI.BankName"
    Expression ="BankPAI.BankAddress"
    Alias ="BankRemitAddress"
    Expression ="[LRemit].[LocationStreet] & Chr(13) & Chr(10) & [LRemit].[LAddrLine1] & Chr(13) "
        "& Chr(10) & IIf(Len([LRemit].[LAddrLine2])>3,[LRemit].[LAddrLine2] & Chr(13) & C"
        "hr(10),\"\") & [LRemit].[LCity] & \" \" & [LRemit].[LState] & \" \" & Left([LRem"
        "it].[LZip],5) & IIf(Len([lremit].[lzip])>5,\"-\" & Right([lremit].[Lzip],4),\"\""
        ")"
    Alias ="Contract"
    Expression ="IIf(Nz([SchBnkClNO],\"\")<>\"\",[SchBnkClNO] & \"-\",\"\") & [SgrpBnkNo]"
    Alias ="GroupsIn"
    Expression ="GroupsNoinSchedule([schedule].[schid],0,False,True)"
    Alias ="ScheduleDate"
    Expression ="Format([SchedDt],\"mmmm d\"\", \"\"yyyy\")"
    Expression ="Clients.clientID"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SgrpId"
    Expression ="Banks.BankID"
End
Begin Joins
    LeftTable ="Banks"
    RightTable ="LRemit"
    Expression ="Banks.BankId = LRemit.ClientREF"
    Flag =2
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Banks"
    Expression ="SchGrp.AssigneeBankID = Banks.BankID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Banks"
    RightTable ="BankPAI"
    Expression ="Banks.BankID = BankPAI.BankID"
    Flag =1
End
Begin OrderBy
    Expression ="SchGrp.SGrpNO"
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
dbInteger "RowHeight" ="1440"
Begin
    Begin
        dbText "Name" ="GroupsIn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="BankRemitAddress"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Contract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentDueDates"
        dbInteger "ColumnWidth" ="3315"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PayDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankPAI.BankAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankPAI.BankName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankPAI.AccountNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankPAI.ABANumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankPAI.AccountName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.blcd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-993
    Top =211
    Right =1139
    Bottom =1202
    Left =-1
    Top =-1
    Right =2108
    Bottom =122
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =604
        Top =15
        Right =796
        Bottom =527
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =435
        Top =23
        Right =579
        Bottom =167
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =828
        Top =93
        Right =1050
        Bottom =482
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =1109
        Top =229
        Right =1253
        Bottom =373
        Top =0
        Name ="LRemit"
        Name =""
    End
    Begin
        Left =22
        Top =19
        Right =165
        Bottom =299
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =241
        Top =35
        Right =385
        Bottom =358
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =1027
        Top =15
        Right =1171
        Bottom =159
        Top =0
        Name ="BankPAI"
        Name =""
    End
End

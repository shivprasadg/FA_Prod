Operation =1
Option =0
Where ="(((ATF.ATFID) Is Not Null))"
Begin InputTables
    Name ="ATF"
    Name ="Clients"
    Name ="Clients"
    Alias ="Clients_1"
End
Begin OutputColumns
    Expression ="ATF.*"
    Alias ="ATFDesc"
    Expression ="\"FAID: \" & [FAID] & \" LPAF Date: \" & [LPAFDate]"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="CoLessee"
    Expression ="Clients_1.clientCompanyName"
End
Begin Joins
    LeftTable ="ATF"
    RightTable ="Clients"
    Expression ="ATF.LesseeID = Clients.clientID"
    Flag =1
    LeftTable ="ATF"
    RightTable ="Clients_1"
    Expression ="ATF.CoLesseeID = Clients_1.clientID"
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
        dbText "Name" ="ATF.SignOffOpsID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.NegBLCD_YN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.CreditRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.TS_Last"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.SignOffPricingDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.PerDiemYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.IsExchangeIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ATFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.SignOffPricingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ProjectedAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ATFDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.DepotContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.LPAFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.SignOffFundingDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.InterimFundingBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.SignOffFundingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.Admin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.LesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.CoLesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.GuarantorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.LeaseTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.NumOfSchedules"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.NumOfGroups"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.NumOfUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.NumOfChildParts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.TitleContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.PurchaseOptionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.SignOffOpsDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.CorpContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ScheduleSignerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.COASignerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CoLessee"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =115
    Top =193
    Right =1714
    Bottom =689
    Left =-1
    Top =-1
    Right =1581
    Bottom =215
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =103
        Top =-2
        Right =259
        Bottom =433
        Top =0
        Name ="ATF"
        Name =""
    End
    Begin
        Left =328
        Top =243
        Right =537
        Bottom =474
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =484
        Top =78
        Right =628
        Bottom =222
        Top =0
        Name ="Clients_1"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="vw_Syndication_Validate_Final"
End
Begin OutputColumns
    Expression ="vw_Syndication_Validate_Final.[Fund Status]"
    Expression ="vw_Syndication_Validate_Final.[Unit Status]"
    Expression ="vw_Syndication_Validate_Final.CompanyName"
    Alias ="MLNo Sch Grp"
    Expression ="[MLno] & \" - \" & [Schedule] & \" - \" & [UnitGroup]"
    Expression ="vw_Syndication_Validate_Final.[FA Year]"
    Expression ="vw_Syndication_Validate_Final.UnitCount"
    Alias ="Proj BLCD"
    Expression ="vw_Syndication_Validate_Final.BLCDProj"
    Expression ="vw_Syndication_Validate_Final.Assignee"
    Expression ="vw_Syndication_Validate_Final.AssetID"
    Expression ="vw_Syndication_Validate_Final.MYear"
    Expression ="vw_Syndication_Validate_Final.Make"
    Expression ="vw_Syndication_Validate_Final.Type"
    Expression ="vw_Syndication_Validate_Final.[Unit Cost]"
    Expression ="vw_Syndication_Validate_Final.[Fund Date (Synd)]"
    Expression ="vw_Syndication_Validate_Final.[Synd_Fee (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Synd_Fee (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Strip Rent (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Strip Rent (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Term-Penalty (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Term-Penalty (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Per Diem]"
    Expression ="vw_Syndication_Validate_Final.[Per Diem (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[UTA at Synd]"
    Expression ="vw_Syndication_Validate_Final.[UTA at Synd (Pct)]"
    Alias ="PV at SYND"
    Expression ="IIf([PVFeesYN]=False,0,PV(([PVRate]/12),[FeeTerm],[FMSS],0,[PVType])*-1)"
    Alias ="PV (Pct)"
    Expression ="IIf([PV at SYND]>0,[PV at SYND]/[Unit Cost],0)"
    Alias ="Interim Expense"
    Expression ="IIf([IntrimLoanDays]>0,[UnitIFValue]+FV(([InterimLoanRate]/360),[IntrimLoanDays]"
        ",0,[UnitIFValue],1),0)"
    Alias ="Interim (Pct)"
    Expression ="IIf([Interim Expense]<>0,[Interim Expense]/[Unit Cost],0)"
    Expression ="vw_Syndication_Validate_Final.Misc_1"
    Expression ="vw_Syndication_Validate_Final.[Misc_1 (Pct)]"
    Expression ="vw_Syndication_Validate_Final.Misc_2"
    Expression ="vw_Syndication_Validate_Final.[Misc_2 (Pct)]"
    Alias ="Syndication Total"
    Expression ="[Syndication_Total]-Abs([Interim Expense])+[PV at SYND]"
    Alias ="Synd Total (Pct)"
    Expression ="([Syndication_Total]-Abs([Interim Expense])+[PV at SYND])/[Unit Cost]"
    Expression ="vw_Syndication_Validate_Final.[Fund Date (LO)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Fee (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Fee (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Term-Penalty (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Term-Penalty (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Per Diem (Avg)]"
    Expression ="vw_Syndication_Validate_Final.[Proj Per Diem (Pct)]"
    Expression ="vw_Syndication_Validate_Final.[Proj UTA]"
    Expression ="vw_Syndication_Validate_Final.[Proj UTA (Pct)]"
    Alias ="Proj PV"
    Expression ="PV(([PVRate]/12),[FeeTerm],[Proj FMSS],0,[PVType])*-1"
    Alias ="Proj PV (Pct)"
    Expression ="IIf([Proj PV]>0,[Proj PV]/[Unit Cost],0)"
    Alias ="Projected Total"
    Expression ="[Projected_Total]+[Proj PV]"
    Alias ="Proj Total (Pct)"
    Expression ="([Projected_Total]+[Proj PV])/[Unit Cost]"
    Expression ="vw_Syndication_Validate_Final.LeaseType"
    Expression ="vw_Syndication_Validate_Final.PVType"
    Expression ="vw_Syndication_Validate_Final.PVRate"
    Alias ="Interim Loan Rate"
    Expression ="vw_Syndication_Validate_Final.InterimLoanRate"
    Expression ="vw_Syndication_Validate_Final.FeeTerm"
    Expression ="vw_Syndication_Validate_Final.SyndNote"
    Expression ="vw_Syndication_Validate_Final.SchRent"
    Expression ="vw_Syndication_Validate_Final.SyndRent"
End
Begin OrderBy
    Expression ="vw_Syndication_Validate_Final.CompanyName"
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
        dbText "Name" ="vw_Syndication_Validate_Final.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[BLCD Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Fund Status]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Unit Status]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Synd_Fee (Avg)]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Term-Penalty (Avg)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Per Diem]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[UTA at Synd]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[UTA at Synd (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj UTA (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.FeeTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Unit Cost]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Synd_Fee (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Fee (Avg)]"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Per Diem (Avg)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj UTA]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.PVType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Strip Rent (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Per Diem (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Term-Penalty (Avg)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Per Diem (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Strip Rent (Avg)]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Term-Penalty (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Fee (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Proj Term-Penalty (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PV at SYND"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PV (Pct)"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj PV (Pct)"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj PV"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.CompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Syndication Total"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Synd Total (Pct)"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Synd Total (Pct) "
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Misc_1 (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.Misc_2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.Misc_1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Misc_2 (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Projected Total"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLNo Sch Grp"
        dbInteger "ColumnWidth" ="4560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Interim (Pct)"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Interim (Pct)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Interim Expense"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.PVRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Interim Loan Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Total (Pct)"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Fund Date (Synd)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[Fund Date (LO)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.[FA Year]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.SyndNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj_BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Syndication_Validate_Final.SyndRent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-382
    Top =55
    Right =1120
    Bottom =1026
    Left =-1
    Top =-1
    Right =1484
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =398
        Top =10
        Right =639
        Bottom =207
        Top =0
        Name ="vw_Syndication_Validate_Final"
        Name =""
    End
End

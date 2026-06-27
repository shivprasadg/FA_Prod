Operation =1
Option =0
Having ="(((Units.DrawDownId)=[Forms]![DrawDownManager]![DrawID]))"
Begin InputTables
    Name ="DrawDowns"
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="LastMilesReportedAll"
    Name ="Banks"
    Name ="MstrLease"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Banks.BankId"
    Expression ="Units.DrawDownId"
    Expression ="vw_SixKeys.clientID"
    Expression ="MstrLease.MLGuarantor"
    Expression ="DrawDowns.DrawNumber"
    Expression ="Banks.BankName"
    Alias ="Guaranty"
    Expression ="Clients.clientCompanyName"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="QTY"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="UnitDesc"
    Expression ="vw_SixKeys.PrintedDocs"
    Alias ="UnitCost"
    Expression ="Units.UnitIfAmt"
    Alias ="TotalCost"
    Expression ="Sum(Units.UnitIfAmt)"
    Expression ="LastMilesReportedAll.MPYAllowed"
    Expression ="vw_SixKeys.LeaseTermPrimary"
    Alias ="EstDelivery"
    Expression ="Format([UnitEstDelDate],\"mmm-yyyy\")"
    Expression ="Units.UnitsRent"
    Alias ="RentTotal"
    Expression ="Sum(Units.[UnitsRent])"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="vw_SixKeys"
    Expression ="DrawDowns.DrawDownId = vw_SixKeys.DrawDownId"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="LastMilesReportedAll"
    Expression ="Units.UnitID = LastMilesReportedAll.UnitID"
    Flag =2
    LeftTable ="DrawDowns"
    RightTable ="Banks"
    Expression ="DrawDowns.DrawBankId = Banks.BankId"
    Flag =1
    LeftTable ="DrawDowns"
    RightTable ="MstrLease"
    Expression ="DrawDowns.MLID = MstrLease.MLID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Clients"
    Expression ="MstrLease.MLGuarantor = Clients.clientID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.UnitGroup"
    Flag =0
    Expression ="Format([UnitEstDelDate],\"mmm-yyyy\")"
    Flag =0
End
Begin Groups
    Expression ="Banks.BankId"
    GroupLevel =0
    Expression ="Units.DrawDownId"
    GroupLevel =0
    Expression ="vw_SixKeys.clientID"
    GroupLevel =0
    Expression ="MstrLease.MLGuarantor"
    GroupLevel =0
    Expression ="DrawDowns.DrawNumber"
    GroupLevel =0
    Expression ="Banks.BankName"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyName"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="vw_SixKeys.PrintedDocs"
    GroupLevel =0
    Expression ="Units.UnitIfAmt"
    GroupLevel =0
    Expression ="LastMilesReportedAll.MPYAllowed"
    GroupLevel =0
    Expression ="vw_SixKeys.LeaseTermPrimary"
    GroupLevel =0
    Expression ="Format([UnitEstDelDate],\"mmm-yyyy\")"
    GroupLevel =0
    Expression ="Units.UnitsRent"
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
Begin
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbInteger "ColumnWidth" ="3765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastMilesReportedAll.MPYAllowed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarantor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentTotal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QTY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="5970"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankName"
        dbInteger "ColumnWidth" ="3765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Guaranty"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstDelivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-618
    Top =194
    Right =1127
    Bottom =823
    Left =-1
    Top =-1
    Right =1721
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =33
        Top =24
        Right =259
        Bottom =411
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =291
        Top =137
        Right =491
        Bottom =650
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =1105
        Top =27
        Right =1366
        Bottom =428
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =855
        Top =340
        Right =1043
        Bottom =664
        Top =0
        Name ="LastMilesReportedAll"
        Name =""
    End
    Begin
        Left =33
        Top =422
        Right =256
        Bottom =566
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =603
        Top =44
        Right =759
        Bottom =224
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =853
        Top =39
        Right =1043
        Bottom =230
        Top =0
        Name ="Clients"
        Name =""
    End
End

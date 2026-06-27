Operation =1
Option =0
Where ="(((vw_SixKeys.PortfolioNumber) Is Not Null) AND ((vw_SixKeys.clientID) Is Not Nu"
    "ll))"
Begin InputTables
    Name ="Schedule_Group_WorkTable"
    Name ="vw_SixKeys"
    Name ="TitleSentTrustDate"
    Name ="Units"
    Name ="Locations"
    Name ="Banks"
    Name ="MstrLease"
End
Begin OutputColumns
    Alias ="Lessee"
    Expression ="vw_SixKeys.CompanyName"
    Alias ="Leasee Master Lease Number"
    Expression ="vw_SixKeys.MLNo"
    Alias ="Leasee Schedule No"
    Expression ="vw_SixKeys.Schedule"
    Alias ="File Type"
    Expression ="\"Equipment File\""
    Alias ="DateSent"
    Expression ="Date()"
    Alias ="TitleNum"
    Expression ="Units.UnitTitleNumber"
    Alias ="State of Title"
    Expression ="Locations.LState"
    Expression ="vw_SixKeys.VIN"
    Alias ="Equipment Description"
    Expression ="[UnitYr] & \"  \" & [UnitMake] & \" \" & [unitmodel] & \" \" & [UnitType] & \" \""
        " & [Axle]"
    Alias ="Portfolio Certificate No"
    Expression ="vw_SixKeys.PortfolioNumber"
    Alias ="Lender or Holder Name"
    Expression ="Banks.BankName"
    Alias ="Copy of Certificate of Insurance included"
    Expression ="IIf([MLLiabIns]=\"R\" And [MlinsExp]>Date(),\"Yes\",\"Yes\")"
    Alias ="Original Bill of Sale Included"
    Expression ="\"N/A\""
    Alias ="FA-Asset Num"
    Expression ="vw_SixKeys.AssetID"
    Expression ="Banks.BankId"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="TitleSentTrustDate"
    Expression ="vw_SixKeys.AssetID = TitleSentTrustDate.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.AssetID = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitTitleLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Banks"
    Expression ="vw_SixKeys.BankId = Banks.BankId"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="MstrLease"
    Expression ="vw_SixKeys.MLID = MstrLease.MLID"
    Flag =1
    LeftTable ="Schedule_Group_WorkTable"
    RightTable ="vw_SixKeys"
    Expression ="Schedule_Group_WorkTable.ScheduleId = vw_SixKeys.SchID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
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
        dbText "Name" ="Original Bill of Sale Included"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leasee Master Lease Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leasee Schedule No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Equipment Description"
        dbInteger "ColumnWidth" ="3480"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA-Asset Num"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lender or Holder Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Portfolio Certificate No"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="File Type"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State of Title"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleNum"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Copy of Certificate of Insurance included"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateSent"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =195
    Top =68
    Right =2653
    Bottom =786
    Left =-1
    Top =-1
    Right =2440
    Bottom =316
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =7
        Top =5
        Right =151
        Bottom =149
        Top =0
        Name ="Schedule_Group_WorkTable"
        Name =""
    End
    Begin
        Left =199
        Top =12
        Right =343
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =764
        Top =76
        Right =908
        Bottom =220
        Top =0
        Name ="TitleSentTrustDate"
        Name =""
    End
    Begin
        Left =415
        Top =133
        Right =612
        Bottom =526
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =726
        Top =244
        Right =870
        Bottom =360
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =914
        Top =68
        Right =1067
        Bottom =397
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =1116
        Top =50
        Right =1279
        Bottom =428
        Top =0
        Name ="MstrLease"
        Name =""
    End
End

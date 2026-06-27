Operation =1
Option =2
Having ="(((SchGrp.POVendor) Is Not Null))"
Begin InputTables
    Name ="Tracking_Open_Units_Clients"
    Name ="SchGrp"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="Tracking_Open_Units_Clients.FAID"
    Expression ="Tracking_Open_Units_Clients.VendorID"
    Expression ="Tracking_Open_Units_Clients.Client"
    Alias ="ClientName"
    Expression ="Tracking_Open_Units_Clients.clientCompanyName"
    Expression ="SchGrp.SGrpNO"
    Alias ="PO"
    Expression ="Nz([PONum1],\"TBD\")"
    Expression ="SchGrp.POVendor"
    Expression ="Vendors.VendorName"
    Expression ="Tracking_Open_Units_Clients.Units"
End
Begin Joins
    LeftTable ="Tracking_Open_Units_Clients"
    RightTable ="SchGrp"
    Expression ="Tracking_Open_Units_Clients.FAID = SchGrp.SGrpFAID"
    Flag =1
    LeftTable ="Tracking_Open_Units_Clients"
    RightTable ="Vendors"
    Expression ="Tracking_Open_Units_Clients.VendorID = Vendors.VendorID"
    Flag =1
End
Begin Groups
    Expression ="Tracking_Open_Units_Clients.FAID"
    GroupLevel =0
    Expression ="Tracking_Open_Units_Clients.VendorID"
    GroupLevel =0
    Expression ="Tracking_Open_Units_Clients.Client"
    GroupLevel =0
    Expression ="Tracking_Open_Units_Clients.clientCompanyName"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Nz([PONum1],\"TBD\")"
    GroupLevel =0
    Expression ="SchGrp.POVendor"
    GroupLevel =0
    Expression ="Vendors.VendorName"
    GroupLevel =0
    Expression ="Tracking_Open_Units_Clients.Units"
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
        dbText "Name" ="Tracking_Open_Units_Clients.Client"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PO"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracking_Open_Units_Clients.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracking_Open_Units_Clients.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="3570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracking_Open_Units_Clients.Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientName"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1507
    Bottom =855
    Left =-1
    Top =-1
    Right =1475
    Bottom =430
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =109
        Top =70
        Right =286
        Bottom =342
        Top =0
        Name ="Tracking_Open_Units_Clients"
        Name =""
    End
    Begin
        Left =357
        Top =16
        Right =560
        Bottom =455
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =686
        Top =161
        Right =830
        Bottom =305
        Top =0
        Name ="Vendors"
        Name =""
    End
End

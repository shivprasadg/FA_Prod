Operation =1
Option =0
Where ="(((Vendors.VendorName)<>\"Lessee\" And (Vendors.VendorName)<>\"unknown\"))"
Begin InputTables
    Name ="Vendors"
    Name ="INV_MsoContactq"
End
Begin OutputColumns
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorID"
    Expression ="INV_MsoContactq.contactID"
    Expression ="INV_MsoContactq.VendorRef"
    Expression ="INV_MsoContactq.contactEmail"
    Expression ="INV_MsoContactq.VendorContact"
End
Begin Joins
    LeftTable ="Vendors"
    RightTable ="INV_MsoContactq"
    Expression ="Vendors.VendorID = INV_MsoContactq.VendorRef"
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
        dbText "Name" ="Vendors.VendorName"
        dbInteger "ColumnWidth" ="4695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INV_MsoContactq.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INV_MsoContactq.VendorRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INV_MsoContactq.contactEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INV_MsoContactq.VendorContact"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =390
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="INV_MsoContactq"
        Name =""
    End
End

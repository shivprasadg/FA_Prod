Operation =1
Option =0
Begin InputTables
    Name ="VendorContacts"
    Name ="Vendors"
End
Begin OutputColumns
End
Begin Joins
    LeftTable ="Vendors"
    RightTable ="VendorContacts"
    Expression ="Vendors.VendorID = VendorContacts.VendorRef"
    Flag =1
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
    Bottom =299
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =244
        Top =11
        Right =388
        Bottom =155
        Top =0
        Name ="VendorContacts"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Vendors"
        Name =""
    End
End

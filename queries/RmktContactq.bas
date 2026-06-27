Operation =1
Option =0
Where ="(((Locations.ClientREF)=forms!RemarketingClientView!clientid) And ((Locations.Lo"
    "cObsolete)=No))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.LocationID"
    Expression ="Locations.LocationName"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Locations.LocationID"
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
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End

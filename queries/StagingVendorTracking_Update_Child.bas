Operation =1
Option =0
Begin InputTables
    Name ="StagingVendorTracking"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="StagingVendorTracking.ClientId"
    Expression ="ChildParts.REFENG"
    Expression ="ChildParts.COMPRESSS"
    Expression ="ChildParts.MICROPROC"
    Expression ="ChildParts.ELECDISP"
End
Begin Joins
    LeftTable ="StagingVendorTracking"
    RightTable ="ChildParts"
    Expression ="StagingVendorTracking.ClientId = ChildParts.ChildPartID"
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
    Begin
        dbText "Name" ="ChildParts.ELECDISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ClientId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =152
    Top =224
    Right =1251
    Bottom =808
    Left =-1
    Top =-1
    Right =1067
    Bottom =284
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =7
        Top =10
        Right =229
        Bottom =302
        Top =0
        Name ="StagingVendorTracking"
        Name =""
    End
    Begin
        Left =272
        Top =19
        Right =495
        Bottom =301
        Top =0
        Name ="ChildParts"
        Name =""
    End
End

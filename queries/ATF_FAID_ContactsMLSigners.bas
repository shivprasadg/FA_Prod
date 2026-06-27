Operation =1
Option =1
Where ="SchID=2958"
Begin InputTables
    Name ="ATF_FAID_ContactsMLSignersBase"
End
Begin OutputColumns
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
dbMemo "Filter" ="([ATF_FAID_MLSigners].[SchID]=251)"
Begin
End
Begin
    State =0
    Left =27
    Top =53
    Right =1827
    Bottom =261
    Left =-1
    Top =-1
    Right =1782
    Bottom =150
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =414
        Bottom =421
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =447
        Top =91
        Right =591
        Bottom =235
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =639
        Top =12
        Right =783
        Bottom =156
        Top =0
        Name ="COA_Signers_Unit"
        Name =""
    End
End

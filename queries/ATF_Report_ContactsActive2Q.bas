Operation =1
Option =1
Where ="SchID=1092"
Begin InputTables
    Name ="ATF_Report_ContactsActive2"
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
dbMemo "Filter" ="([ATF_Report_ContactsActive2].[SchID]=932)"
Begin
End
Begin
    State =0
    Left =94
    Top =195
    Right =1571
    Bottom =851
    Left =-1
    Top =-1
    Right =1459
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =221
        Top =20
        Right =365
        Bottom =164
        Top =0
        Name ="ATF_ContactsActive"
        Name =""
    End
    Begin
        Left =31
        Top =20
        Right =175
        Bottom =164
        Top =0
        Name ="ATF_Report_Active_ContactControl"
        Name =""
    End
End

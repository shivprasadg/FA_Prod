Operation =1
Option =0
Where ="(((vw_SixKeys.Status) Like \"*S*\" Or (vw_SixKeys.Status) Like \"*K*\" Or (vw_Si"
    "xKeys.Status) Like '*T*' Or (vw_SixKeys.Status) Like '*R*'))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.*"
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
    Left =50
    Top =84
    Right =1350
    Bottom =807
    Left =-1
    Top =-1
    Right =1282
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =84
        Top =3
        Right =496
        Bottom =388
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End

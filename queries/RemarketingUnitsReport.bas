Operation =1
Option =0
Where ="(((Rmkt.InvoicePaidDate) Is Null Or (Rmkt.InvoicePaidDate) Between [forms]![rmkt"
    " Menu]![FromDate] And [forms]![rmkt Menu]![ToDate]))"
Begin InputTables
    Name ="Rmkt"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
End
Begin Joins
    LeftTable ="Rmkt"
    RightTable ="vw_SixKeys"
    Expression ="Rmkt.UnitRef = vw_SixKeys.AssetID"
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
    Left =412
    Top =192
    Right =1695
    Bottom =915
    Left =-1
    Top =-1
    Right =1251
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =42
        Top =38
        Right =241
        Bottom =372
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =262
        Top =42
        Right =460
        Bottom =441
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End

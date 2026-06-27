Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="YearInServiced"
    Expression ="IIf([Status]=\"B\",2016,DatePart(\"yyyy\",[InServiceDate]))"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitID)"
End
Begin OrderBy
    Expression ="IIf([Status]=\"B\",2016,DatePart(\"yyyy\",[InServiceDate]))"
    Flag =1
End
Begin Groups
    Expression ="IIf([Status]=\"B\",2016,DatePart(\"yyyy\",[InServiceDate]))"
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
        dbText "Name" ="YearInServiced"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =162
    Right =1344
    Bottom =1043
    Left =-1
    Top =-1
    Right =1270
    Bottom =580
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =235
        Top =44
        Right =436
        Bottom =413
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End

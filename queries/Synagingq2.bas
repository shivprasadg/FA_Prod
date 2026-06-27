Operation =1
Option =0
Where ="(((Synagingq1.UnitsInGroup)=[CountOfunitvin]) AND ((Clients.ClientGroupID) Not L"
    "ike \"* demo\"))"
Begin InputTables
    Name ="Clients"
    Name ="Synagingq1"
    Name ="Synagingq1a"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Synagingq1.ClientShNm"
    Expression ="Synagingq1.MLNo"
    Alias ="Sch No"
    Expression ="Synagingq1.SchNo"
    Alias ="Grp No"
    Expression ="Synagingq1.SGrpNO"
    Alias ="BLCD"
    Expression ="Synagingq1.BLCD"
    Alias ="Bank Assign Dt"
    Expression ="Synagingq1.AssignmentDate"
    Alias ="No Of Units"
    Expression ="Synagingq1.UnitsInGroup"
    Alias ="Units In-Serv"
    Expression ="IIf(IsNull([INserVIN]),0,[inserVIN])"
    Alias ="Max Accepteded Dt"
    Expression ="Synagingq1.MaxOfunitacceptdate"
    Alias ="Age Accepted"
    Expression ="Date()-[maxofunitacceptdate]+1"
    Alias ="AccpetDt+90"
    Expression ="DateAdd(\"d\",90,[maxofunitacceptdate])"
    Alias ="Max Inservice dt"
    Expression ="IIf(IsNull([inservin]) Or [UnitsInGroup]<>[inservin],\"TBD\",[maxofunitinserv])"
    Alias ="Age In-Service"
    Expression ="Date()-[maxofunitinserv]+1"
    Alias ="In-Inserv+90"
    Expression ="DateAdd(\"d\",90,[maxofunitinserv])"
    Alias ="As of"
    Expression ="Date()"
    Expression ="Clients.ClientGroupID"
End
Begin Joins
    LeftTable ="Synagingq1"
    RightTable ="Synagingq1a"
    Expression ="Synagingq1.SchID = Synagingq1a.SchID"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Synagingq1"
    Expression ="Clients.ClientShNm = Synagingq1.ClientShNm"
    Flag =1
End
Begin OrderBy
    Expression ="Date()-[maxofunitinserv]+1"
    Flag =1
    Expression ="Synagingq1.ClientShNm"
    Flag =0
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
        dbText "Name" ="Synagingq1.MLNo"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Grp No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Sch No"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="No Of Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Max Accepteded Dt"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Bank Assign Dt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Max Inservice dt"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Age Accepted"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Age In-Service"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units In-Serv"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="AccpetDt+90"
        dbInteger "ColumnWidth" ="1575"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
    End
    Begin
        dbText "Name" ="In-Inserv+90"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
    End
    Begin
        dbText "Name" ="AS of"
        dbText "Format" ="mm/yy/dd"
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
    Bottom =300
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =513
        Top =12
        Right =657
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =273
        Bottom =302
        Top =0
        Name ="Synagingq1"
        Name =""
    End
    Begin
        Left =321
        Top =12
        Right =465
        Bottom =156
        Top =0
        Name ="Synagingq1a"
        Name =""
    End
End

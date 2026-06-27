Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((vw_Units_AutoHaulerPaired.TractorUnitID) Is "
    "Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Clients"
    Name ="EXchangeIT_ByYear"
    Name ="Extended_ByYear"
    Name ="vw_Units_AutoHaulerPaired"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="BecameClient"
    Expression ="Clients.clientBecameDate"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Status"
    Alias ="XITYear"
    Expression ="IIf([OptionTypeID]=1,CInt(DatePart(\"yyyy\",CDate(Nz(Nz([schgrp].[BLCD],Nz([Acce"
        "ptedDate],[ActualDeliveryDate])),Date())))))"
    Alias ="EXTendYear"
    Expression ="Nz([Extended_ByYear].[EXTendYear],0)"
    Alias ="BCYear"
    Expression ="IIf(Abs(DateDiff(\"m\",[clientBecameDate],[schgrp].[BLCD]))<=12,[Dyear])"
    Alias ="DYear"
    Expression ="IIf([status]=\"B\",CDate(\"1/31/2020\"),CInt(DatePart(\"yyyy\",CDate(Nz(Nz([fund"
        "date],Nz([schgrp].[BLCD],Nz([AcceptedDate],[ActualDeliveryDate]))),\"1/31/2020\""
        ")))))"
    Alias ="OYear"
    Expression ="CInt(DatePart(\"yyyy\",CDate(Nz([OffLeaseDate],Nz([ProjEOL],\"12/31/2050\")))))"
    Alias ="A2012"
    Expression ="IIf([Dyear]<=2012 And [OYear]>2012,1,0)"
    Alias ="A2013"
    Expression ="IIf([Dyear]<=2013 And [OYear]>2013,1,0)"
    Alias ="A2014"
    Expression ="IIf([Dyear]<=2014 And [OYear]>2014,1,0)"
    Alias ="A2015"
    Expression ="IIf([Dyear]<=2015 And [OYear]>2015,1,0)"
    Alias ="A2016"
    Expression ="IIf([Dyear]<=2016 And [OYear]>2016,1,0)"
    Alias ="A2017"
    Expression ="IIf([Dyear]<=2017 And [OYear]>2017,1,0)"
    Alias ="A2018"
    Expression ="IIf([Dyear]<=2018 And [OYear]>2018,1,0)"
    Alias ="D2013"
    Expression ="IIf([DYear]=2013,1,0)"
    Alias ="D2014"
    Expression ="IIf([DYear]=2014,1,0)"
    Alias ="D2015"
    Expression ="IIf([DYear]=2015,1,0)"
    Alias ="D2016"
    Expression ="IIf([DYear]=2016,1,0)"
    Alias ="D2017"
    Expression ="IIf([DYear]=2017,1,0)"
    Alias ="D2018"
    Expression ="IIf([DYear]=2018,1,0)"
    Alias ="O2013"
    Expression ="IIf([OYear]=2013,1,0)"
    Alias ="O2014"
    Expression ="IIf([OYear]=2014,1,0)"
    Alias ="O2015"
    Expression ="IIf([OYear]=2015,1,0)"
    Alias ="O2016"
    Expression ="IIf([OYear]=2016,1,0)"
    Alias ="O2017"
    Expression ="IIf([OYear]=2017,1,0)"
    Alias ="O2018"
    Expression ="IIf([OYear]=2018,1,0)"
    Alias ="O2019"
    Expression ="IIf([OYear]=2019,1,0)"
    Alias ="N2013"
    Expression ="IIf([BCYear]=2013,1,0)"
    Alias ="N2014"
    Expression ="IIf([BCYear]=2014,1,0)"
    Alias ="N2015"
    Expression ="IIf([BCYear]=2015,1,0)"
    Alias ="N2016"
    Expression ="IIf([BCYear]=2016,1,0)"
    Alias ="N2017"
    Expression ="IIf([BCYear]=2017,1,0)"
    Alias ="N2018"
    Expression ="IIf([BCYear]=2018,1,0)"
    Alias ="XIT2013"
    Expression ="0"
    Alias ="XIT2014"
    Expression ="IIf([XITYear]=2014,1,0)"
    Alias ="XIT2015"
    Expression ="IIf([XITYear]=2015,1,0)"
    Alias ="XIT2016"
    Expression ="IIf([XITYear]=2016,1,0)"
    Alias ="XIT2017"
    Expression ="IIf([XITYear]=2017,1,0)"
    Alias ="XIT2018"
    Expression ="IIf([XITYear]=2018,1,0)"
    Alias ="EYear2013"
    Expression ="0"
    Alias ="EYear2014"
    Expression ="0"
    Alias ="EYear2015"
    Expression ="0"
    Alias ="EYear2016"
    Expression ="0"
    Alias ="EYear2017"
    Expression ="IIf([EXTendYear]=2017,1,0)"
    Alias ="EYear2018"
    Expression ="IIf([EXTendYear]=2018,1,0)"
    Alias ="B2018"
    Expression ="IIf([Status]=\"B\",1,0)"
    Alias ="TR2018"
    Expression ="IIf([Status]=\"TR\",1,0)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Clients"
    Expression ="vw_SixKeys.clientID = Clients.clientID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="EXchangeIT_ByYear"
    Expression ="vw_SixKeys.SGrpID = EXchangeIT_ByYear.ReplacedByGroupID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Extended_ByYear"
    Expression ="vw_SixKeys.AssetID = Extended_ByYear.UnitID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="vw_Units_AutoHaulerPaired"
    Expression ="vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
End
Begin OrderBy
    Expression ="IIf([status]=\"B\",CDate(\"1/31/2020\"),CInt(DatePart(\"yyyy\",CDate(Nz(Nz([fund"
        "date],Nz([schgrp].[BLCD],Nz([AcceptedDate],[ActualDeliveryDate]))),\"1/31/2020\""
        ")))))"
    Flag =0
    Expression ="vw_SixKeys.ClientShortName"
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
dbMemo "Filter" ="([ActiveUnits_ByYear].[ClientShortName] Like \"*toyo*\")"
Begin
    Begin
        dbText "Name" ="DYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BecameClient"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BCYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XITYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXTendYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TR2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XIT2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EYear2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TRYear2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BYear2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired.TractorUnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =131
    Top =63
    Right =1681
    Bottom =1009
    Left =-1
    Top =-1
    Right =1532
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =217
        Bottom =637
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =286
        Top =116
        Right =430
        Bottom =571
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =476
        Top =115
        Right =620
        Bottom =259
        Top =0
        Name ="EXchangeIT_ByYear"
        Name =""
    End
    Begin
        Left =666
        Top =172
        Right =810
        Bottom =375
        Top =0
        Name ="Extended_ByYear"
        Name =""
    End
    Begin
        Left =755
        Top =27
        Right =899
        Bottom =171
        Top =0
        Name ="vw_Units_AutoHaulerPaired"
        Name =""
    End
    Begin
        Left =947
        Top =12
        Right =1091
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Cost_OBC"
    Name ="Cost_OBC_Adapt"
    Name ="Cost_OBC_Dome"
    Name ="Cost_OBC_Multi"
    Name ="Cost_OBC_Install"
    Name ="Cost_OBC_Display"
    Name ="Cost_OBC_Part"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="Units.UnitID"
    Expression ="Cost_OBC.OBC_Make"
    Alias ="OBC_Cost"
    Expression ="CCur(Nz([OBC],0))"
    Expression ="Cost_OBC.OBCInvoice"
    Expression ="Cost_OBC.OBCPaid"
    Expression ="Cost_OBC_Adapt.OBC_Adapt"
    Alias ="OBCAdaptr"
    Expression ="CCur(Nz([OBCAdapt],0))"
    Expression ="Cost_OBC_Adapt.OBCAdaptInv"
    Expression ="Cost_OBC_Adapt.OBCAdaptPaid"
    Expression ="Cost_OBC_Dome.OBC_Dome"
    Alias ="OBCDome"
    Expression ="CCur(Nz([Cost_OBC_Dome].[OBCDome],0))"
    Expression ="Cost_OBC_Dome.OBCDomeInv"
    Expression ="Cost_OBC_Dome.OBCDomePaid"
    Expression ="Cost_OBC_Multi.OBC_Multi"
    Alias ="OBCMulti"
    Expression ="CCur(Nz([Cost_OBC_Multi].[OBCMulti]))"
    Expression ="Cost_OBC_Multi.OBCMultiInv"
    Expression ="Cost_OBC_Multi.OBCMultiPaid"
    Expression ="Cost_OBC_Display.OBC_Display"
    Alias ="OBCDisplay"
    Expression ="CCur(Nz([Cost_OBC_Display].[OBCDisplay],0))"
    Expression ="Cost_OBC_Display.OBCDisplayInv"
    Expression ="Cost_OBC_Display.OBCDisplayPaid"
    Expression ="Cost_OBC_Install.OBCInstaller"
    Alias ="OBCInstall"
    Expression ="CCur(Nz([Cost_OBC_Install].[OBCInstall],0))"
    Expression ="Cost_OBC_Install.OBCInsInvoice"
    Expression ="Cost_OBC_Install.OBCInstallPaid"
    Expression ="Cost_OBC_Part.OBC_Part"
    Alias ="OBCPart"
    Expression ="CCur(Nz([Cost_OBC_Part].[OBCPart],0))"
    Expression ="Cost_OBC_Part.OBCPartInv"
    Expression ="Cost_OBC_Part.OBCPartPaid"
    Alias ="OBCtotal"
    Expression ="CCur(Nz([OBC],0)+Nz([OBCAdapt],0)+Nz([OBCDome],0)+Nz([OBCMulti],0)+Nz([OBCInstal"
        "l],0)+Nz([Cost_OBC_Display].[OBCDisplay],0)+[OBCPart])"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Cost_OBC"
    Expression ="Units.UnitID = Cost_OBC.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Adapt"
    Expression ="Units.UnitID = Cost_OBC_Adapt.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Dome"
    Expression ="Units.UnitID = Cost_OBC_Dome.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Multi"
    Expression ="Units.UnitID = Cost_OBC_Multi.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Install"
    Expression ="Units.UnitID = Cost_OBC_Install.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Display"
    Expression ="Units.UnitID = Cost_OBC_Display.AssetId"
    Flag =2
    LeftTable ="Units"
    RightTable ="Cost_OBC_Part"
    Expression ="Units.UnitID = Cost_OBC_Part.AssetId"
    Flag =2
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
dbMemo "Filter" ="([Cost_OBC_All].[AssetId]=276143)"
Begin
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC.OBCPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Adapt.OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Dome.OBCDomePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Multi.OBCMultiPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Install.OBCInstallPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC.OBC_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC.OBCInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Adapt.OBC_Adapt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Adapt.OBCAdaptInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Dome.OBC_Dome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Dome.OBCDomeInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Multi.OBC_Multi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Multi.OBCMultiInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Cost_OBC_Install.OBCInsInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Install.OBCInstaller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCMulti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCtotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Display.OBCDisplayPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Display.OBCDisplayInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Display.OBC_Display"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Part.OBC_Part"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_Part.OBCPartPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Cost_OBC_Part.OBCPartInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Part"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartPaid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =232
    Top =339
    Right =2541
    Bottom =995
    Left =-1
    Top =-1
    Right =2277
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =17
        Top =10
        Right =158
        Bottom =277
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =176
        Top =14
        Right =320
        Bottom =199
        Top =0
        Name ="Cost_OBC"
        Name =""
    End
    Begin
        Left =330
        Top =43
        Right =474
        Bottom =215
        Top =0
        Name ="Cost_OBC_Adapt"
        Name =""
    End
    Begin
        Left =484
        Top =87
        Right =628
        Bottom =231
        Top =0
        Name ="Cost_OBC_Dome"
        Name =""
    End
    Begin
        Left =642
        Top =123
        Right =786
        Bottom =267
        Top =0
        Name ="Cost_OBC_Multi"
        Name =""
    End
    Begin
        Left =960
        Top =195
        Right =1104
        Bottom =339
        Top =0
        Name ="Cost_OBC_Install"
        Name =""
    End
    Begin
        Left =798
        Top =161
        Right =950
        Bottom =314
        Top =0
        Name ="Cost_OBC_Display"
        Name =""
    End
    Begin
        Left =981
        Top =33
        Right =1125
        Bottom =177
        Top =0
        Name ="Cost_OBC_Part"
        Name =""
    End
End

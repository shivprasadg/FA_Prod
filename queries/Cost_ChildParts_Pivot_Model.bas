Operation =6
Option =0
Begin InputTables
    Name ="Cost_ChildParts"
End
Begin OutputColumns
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Alias ="Expr1"
    Expression ="[PartClass] & \" Type\""
    GroupLevel =1
    Alias ="FirstOfPartModel"
    Expression ="First(Cost_ChildParts.PartType)"
End
Begin Groups
    Expression ="Cost_ChildParts.AssetID"
    GroupLevel =2
    Expression ="[PartClass] & \" Type\""
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Cost_ChildParts.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.PartPaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.PartType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.PartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_Types.PartType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.PartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deer Guards"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Headboards & Toolboxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inverters"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VarOfPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="APU"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Labor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts.ChildVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfChildVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deer GuardsPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvertersPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="APUPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Headboards & ToolboxesPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LaborPartInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Extended Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2yr / 200k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2yr Unlimited Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr Unlimited Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 600k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4yr / 400k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4yr / 450k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4yr / 500k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit 2 Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Van Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="18 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit Other Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Installation Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7yr / 250k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Doc Fee Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="22 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="23 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="24 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="26 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="28 ft Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3yr / 300k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3yr / 350k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3yr / 450k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3yr / 500k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty CARB Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3yr Unlimited Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4yr / 525k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4yr Unlimited Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Special Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6yr Unlimited Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Refrig Body Type"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 100k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 200k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tank Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 250k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Company Logo Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6yr / 750k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 300k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 350k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 400k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 500k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5yr / 550k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6yr / 300k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Processing Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6yr / 600k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7yr / 350k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7yr / 400k Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Attachment Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bendix Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Blower Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Box Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bulk Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cab Logo Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cables / Adapter Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Port Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cables / Dome Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cables / Multiport Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Camara Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Camera Cables Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Collision Avoidance Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Coolant Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Refrig Multi-Temp Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Curtainside Dbl Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dry Body Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cut Away Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tariff 2 Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deer Guard Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delivery Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DOT Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Electric Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Labor Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emissions Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Flap/Fender/Plate Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Floor Mat Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fuel Tank Step Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Headramp Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Heater Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Infotainment Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Knuckle Boom Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Install Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Insulation Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inverter Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lift Gate Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Load Lock carriers Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Conv Body Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Module Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Other Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plate Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pump Wet Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rack Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rigging Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rollup door Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Roof Mod Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Skirt Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC Other Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Snow Plow Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Step Van Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Toolbox Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Surcharge Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="System / Software Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tarp Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tarriff Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBD Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Fee Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Towel Bar labor Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Towel Bar Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TPM Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trailer Connection Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trailer Logo Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Training Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Truck Mount Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit Number Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unkn Cust Paid Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit Kit Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UpFront Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Upgraded Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Various Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wallbox EV Charger Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Accessory Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Add-On Fee Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="APU Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Body Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Boom Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Conveyor Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Crane Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal  Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal Other Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decal Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dump Body Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Electronics Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Flat Bed Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Transmission Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Forklift Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Freight Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maintenance 2 Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maintenance Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Other Body Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paper Work Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTO Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tariff Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tax Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tires Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit Labor Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Up Fit Part Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Walboard Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty ATS Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Chassis / Refrig Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Chassis / Tow Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Eng / Tow Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Chassis Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Clutch Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Eng / ATS / CARB Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Eng / ATS Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Engine Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty HVAC Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty PowerTrain Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Refrig Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty Tow Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-504
    Top =79
    Right =1023
    Bottom =1030
    Left =-1
    Top =-1
    Right =1503
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =54
        Top =63
        Right =223
        Bottom =330
        Top =0
        Name ="Cost_ChildParts"
        Name =""
    End
End

Operation =3
Name ="UnitVinDetails"
Option =0
Where ="(((UnitVINResults.ClientID)=[Ask_ClientID]) AND ((UnitVINResults.make)<>\"\") AN"
    "D ((UnitVinDetails.vin) Is Null))"
Begin InputTables
    Name ="UnitVINResults"
    Name ="UnitVinDetails"
    Name ="Units"
End
Begin OutputColumns
    Name ="ClientID"
    Expression ="UnitVINResults.ClientID"
    Name ="UnitID"
    Expression ="Units.UnitID"
    Name ="vin"
    Expression ="UnitVINResults.vin"
    Name ="modelyear"
    Expression ="UnitVINResults.modelyear"
    Name ="make"
    Expression ="UnitVINResults.make"
    Name ="model"
    Expression ="UnitVINResults.model"
    Name ="trimlevel"
    Expression ="UnitVINResults.trimlevel"
    Name ="enginetype"
    Expression ="UnitVINResults.enginetype"
    Name ="transmission"
    Expression ="UnitVINResults.transmission"
    Name ="drivelinetype"
    Expression ="UnitVINResults.drivelinetype"
    Name ="bodytype"
    Expression ="UnitVINResults.bodytype"
    Name ="vehicletype"
    Expression ="UnitVINResults.vehicletype"
    Name ="vehicleclass"
    Expression ="UnitVINResults.vehicleclass"
    Name ="assyplant"
    Expression ="UnitVINResults.assyplant"
    Name ="cabtype"
    Expression ="UnitVINResults.cabtype"
    Name ="trailertype"
    Expression ="UnitVINResults.trailertype"
    Name ="bustype"
    Expression ="UnitVINResults.bustype"
    Name ="fueltype"
    Expression ="UnitVINResults.fueltype"
    Name ="carlinecode"
    Expression ="UnitVINResults.carlinecode"
    Name ="chassis"
    Expression ="UnitVINResults.chassis"
    Name ="modelchange"
    Expression ="UnitVINResults.modelchange"
    Name ="country"
    Expression ="UnitVINResults.country"
    Name ="motorcycletypes"
    Expression ="UnitVINResults.motorcycletypes"
    Name ="horsepower"
    Expression ="UnitVINResults.horsepower"
    Name ="engineseriescode"
    Expression ="UnitVINResults.engineseriescode"
    Name ="gvwrclass"
    Expression ="UnitVINResults.gvwrclass"
    Name ="gvwrprodrange"
    Expression ="UnitVINResults.gvwrprodrange"
    Name ="wheelbase"
    Expression ="UnitVINResults.wheelbase"
    Name ="axleconfiguration"
    Expression ="UnitVINResults.axleconfiguration"
    Name ="manufacturer"
    Expression ="UnitVINResults.manufacturer"
    Name ="enginemanufacturer"
    Expression ="UnitVINResults.enginemanufacturer"
    Name ="wmiidcode"
    Expression ="UnitVINResults.wmiidcode"
    Name ="plantcode"
    Expression ="UnitVINResults.plantcode"
    Name ="restraintsystem"
    Expression ="UnitVINResults.restraintsystem"
    Name ="series"
    Expression ="UnitVINResults.series"
    Name ="enginemodifier"
    Expression ="UnitVINResults.enginemodifier"
    Name ="steering"
    Expression ="UnitVINResults.steering"
    Name ="emissionscode"
    Expression ="UnitVINResults.emissionscode"
    Name ="brakesystem"
    Expression ="UnitVINResults.brakesystem"
    Name ="frametype"
    Expression ="UnitVINResults.frametype"
    Name ="hybridtype"
    Expression ="UnitVINResults.hybridtype"
    Name ="enginecode"
    Expression ="UnitVINResults.enginecode"
    Name ="status"
    Expression ="UnitVINResults.status"
    Name ="description"
    Expression ="UnitVINResults.description"
End
Begin Joins
    LeftTable ="UnitVINResults"
    RightTable ="UnitVinDetails"
    Expression ="UnitVINResults.vin = UnitVinDetails.vin"
    Flag =2
    LeftTable ="Units"
    RightTable ="UnitVINResults"
    Expression ="Units.UnitVIN = UnitVINResults.VIN"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.vehicletype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.vehicleclass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.assyplant"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.carlinecode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.chassis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.gvwrclass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.plantcode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.series"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.steering"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.cabtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.modelchange"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.country"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.horsepower"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.engineseriescode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.gvwrprodrange"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.wheelbase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.manufacturer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.enginemanufacturer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.restraintsystem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.enginemodifier"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.emissionscode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.brakesystem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.frametype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.unitid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.trimlevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.trailertype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.bustype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.motorcycletypes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.axleconfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.hybridtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.enginecode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDetails.vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.unitid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.enginetype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.transmission"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.drivelinetype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.bodytype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.fueltype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.wmiidcode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.modelyear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVINResults.ClientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =66
    Top =22
    Right =1527
    Bottom =1020
    Left =-1
    Top =-1
    Right =1443
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =313
        Top =94
        Right =543
        Bottom =394
        Top =0
        Name ="UnitVINResults"
        Name =""
    End
    Begin
        Left =677
        Top =129
        Right =821
        Bottom =273
        Top =0
        Name ="UnitVinDetails"
        Name =""
    End
    Begin
        Left =103
        Top =102
        Right =247
        Bottom =519
        Top =0
        Name ="Units"
        Name =""
    End
End

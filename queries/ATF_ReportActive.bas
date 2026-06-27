Operation =1
Option =0
Begin InputTables
    Name ="vw_ATF_Active_Profile"
End
Begin OutputColumns
    Expression ="vw_ATF_Active_Profile.SchID"
    Expression ="vw_ATF_Active_Profile.ATFID"
    Expression ="vw_ATF_Active_Profile.FAID"
    Expression ="vw_ATF_Active_Profile.MLID"
    Expression ="vw_ATF_Active_Profile.LesseeID"
    Expression ="vw_ATF_Active_Profile.CoLesseeID"
    Expression ="vw_ATF_Active_Profile.SubLeasee"
    Expression ="vw_ATF_Active_Profile.CreditRank"
    Expression ="vw_ATF_Active_Profile.ATFSchSpecID"
    Alias ="Assignee_Bank"
    Expression ="Max(vw_ATF_Active_Profile.Assignee_Bank)"
    Expression ="vw_ATF_Active_Profile.SchFAID"
    Expression ="vw_ATF_Active_Profile.CompanyName"
    Expression ="vw_ATF_Active_Profile.ClientShortName"
    Expression ="vw_ATF_Active_Profile.clientTaxID"
    Expression ="vw_ATF_Active_Profile.MLNo"
    Expression ="vw_ATF_Active_Profile.MLOrig"
    Expression ="vw_ATF_Active_Profile.MLDate"
    Alias ="ActualUnits"
    Expression ="Sum(vw_ATF_Active_Profile.ActualUnits)"
    Alias ="SumOfUnitsCost"
    Expression ="Sum(vw_ATF_Active_Profile.UnitsCost)"
    Alias ="SumOfParts"
    Expression ="Sum(vw_ATF_Active_Profile.Parts)"
    Alias ="SumOfPartsCost"
    Expression ="Sum(vw_ATF_Active_Profile.PartsCost)"
    Alias ="SumOfUnitOEC"
    Expression ="Sum(vw_ATF_Active_Profile.UnitOEC)"
    Alias ="AvgOfSchRentAvg"
    Expression ="Avg(vw_ATF_Active_Profile.SchRentAvg)"
    Alias ="AvgOfSyndRentAvg"
    Expression ="Avg(vw_ATF_Active_Profile.SyndRentAvg)"
    Alias ="AvgOfExtendedRentAvg"
    Expression ="Avg(vw_ATF_Active_Profile.ExtendedRentAvg)"
    Alias ="AvgOfMTMRentAvg"
    Expression ="Avg(vw_ATF_Active_Profile.MTMRentAvg)"
    Alias ="AvgOfPDRentAvg"
    Expression ="Avg(vw_ATF_Active_Profile.PDRentAvg)"
    Alias ="AvgOfGroupLeasePayment"
    Expression ="Avg(vw_ATF_Active_Profile.GroupLeasePayment)"
    Alias ="ATFDesc"
    Expression ="vw_ATF_Active_Profile.MLNo"
    Alias ="NumOfSchedules"
    Expression ="Ecount(\"SchID\",\"vw_SixKeys\",\"SchID=\" & [SchID] & \"\",True)"
    Alias ="NumOfGroups"
    Expression ="Ecount(\"sGrpID\",\"vw_SixKeys\",\"SchID=\" & [SchID] & \"\",True)"
    Alias ="NumOfUnits"
    Expression ="Ecount(\"UnitID\",\"vw_SixKeys\",\"SchID=\" & [SchID] & \"\",True)"
    Alias ="CoLessee"
    Expression ="Null"
    Expression ="vw_ATF_Active_Profile.ATFDate"
    Expression ="vw_ATF_Active_Profile.LPAFDate"
    Expression ="vw_ATF_Active_Profile.BDEX"
    Expression ="vw_ATF_Active_Profile.SchDescription"
    Expression ="vw_ATF_Active_Profile.SchNo"
    Expression ="vw_ATF_Active_Profile.Assigned"
    Alias ="SFDCType"
    Expression ="DLookUp(\"SFDCType\",\"Deals\",\"FAID='\" & [FAID] & \"'\")"
    Alias ="GroupsIn"
    Expression ="Count(vw_ATF_Active_Profile.sGrpID)"
    Expression ="vw_ATF_Active_Profile.SGrpID"
    Alias ="GroupID"
    Expression ="vw_ATF_Active_Profile.SGrpID"
    Expression ="vw_ATF_Active_Profile.GroupNo"
End
Begin Groups
    Expression ="vw_ATF_Active_Profile.SchID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.ATFID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.FAID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.MLID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.LesseeID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.CoLesseeID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SubLeasee"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.CreditRank"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.ATFSchSpecID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SchFAID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.CompanyName"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.ClientShortName"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.clientTaxID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.MLOrig"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.MLDate"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.MLNo"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.ATFDate"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.LPAFDate"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.BDEX"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SchDescription"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SchNo"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.Assigned"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SGrpID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.MLNo"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.SGrpID"
    GroupLevel =0
    Expression ="vw_ATF_Active_Profile.GroupNo"
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
        dbText "Name" ="vw_ATF_Active_Profile.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.CoLesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.SubLeasee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.MLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfSchRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.ATFSchSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SFDCType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.LesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.CreditRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfGroupLeasePayment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.clientTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPartsCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfSyndRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitsCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfExtendedRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfParts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfMTMRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfPDRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.SchDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.ATFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumOfSchedules"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumOfGroups"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.Assigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumOfUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CoLessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.LPAFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Active_Profile.GroupNo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =72
    Top =88
    Right =1332
    Bottom =890
    Left =-1
    Top =-1
    Right =1242
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =305
        Bottom =378
        Top =0
        Name ="vw_ATF_Active_Profile"
        Name =""
    End
End

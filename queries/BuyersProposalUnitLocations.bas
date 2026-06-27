Operation =1
Option =0
Begin InputTables
    Name ="Buyer2Proposal_Query"
End
Begin OutputColumns
    Expression ="Buyer2Proposal_Query.LocationID"
    Expression ="Buyer2Proposal_Query.SourceCompany"
    Expression ="Buyer2Proposal_Query.Location"
    Alias ="Units"
    Expression ="Count(Buyer2Proposal_Query.UnitId)"
    Expression ="Buyer2Proposal_Query.ProposalID"
    Expression ="Buyer2Proposal_Query.Latitude"
    Expression ="Buyer2Proposal_Query.Longitude"
End
Begin Groups
    Expression ="Buyer2Proposal_Query.LocationID"
    GroupLevel =0
    Expression ="Buyer2Proposal_Query.SourceCompany"
    GroupLevel =0
    Expression ="Buyer2Proposal_Query.Location"
    GroupLevel =0
    Expression ="Buyer2Proposal_Query.ProposalID"
    GroupLevel =0
    Expression ="Buyer2Proposal_Query.Latitude"
    GroupLevel =0
    Expression ="Buyer2Proposal_Query.Longitude"
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
        dbText "Name" ="Buyer2Proposal_Query.ProposalID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer2Proposal_Query.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer2Proposal_Query.Location"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer2Proposal_Query.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer2Proposal_Query.Longitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer2Proposal_Query.Latitude"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1192
    Bottom =1021
    Left =-1
    Top =-1
    Right =1160
    Bottom =663
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =132
        Top =52
        Right =711
        Bottom =497
        Top =0
        Name ="Buyer2Proposal_Query"
        Name =""
    End
End

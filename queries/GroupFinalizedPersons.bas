Operation =1
Option =0
Begin InputTables
    Name ="GroupFinalizedTrackerFields"
    Name ="DepartmentContacts"
    Name ="TBL_Employees"
End
Begin OutputColumns
    Expression ="GroupFinalizedTrackerFields.FieldName"
    Expression ="GroupFinalizedTrackerFields.FieldLabel"
    Expression ="DepartmentContacts.Department"
    Alias ="Person"
    Expression ="TBL_Employees.[Employee Name]"
    Alias ="email"
    Expression ="[User ID] & \"@fleetadvantage.com\""
End
Begin Joins
    LeftTable ="GroupFinalizedTrackerFields"
    RightTable ="DepartmentContacts"
    Expression ="GroupFinalizedTrackerFields.DepartmentContactID = DepartmentContacts.INDX"
    Flag =1
    LeftTable ="DepartmentContacts"
    RightTable ="TBL_Employees"
    Expression ="DepartmentContacts.PrimaryContactID = TBL_Employees.ID"
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
    Begin
        dbText "Name" ="GroupFinalizedTrackerFields.FieldName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFinalizedTrackerFields.FieldLabel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentContacts.Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="email"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Person"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1307
    Bottom =969
    Left =-1
    Top =-1
    Right =1245
    Bottom =582
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =150
        Top =89
        Right =294
        Bottom =233
        Top =0
        Name ="GroupFinalizedTrackerFields"
        Name =""
    End
    Begin
        Left =342
        Top =12
        Right =486
        Bottom =298
        Top =0
        Name ="DepartmentContacts"
        Name =""
    End
    Begin
        Left =534
        Top =12
        Right =678
        Bottom =156
        Top =0
        Name ="TBL_Employees"
        Name =""
    End
End

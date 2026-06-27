Operation =1
Option =0
Begin InputTables
    Name ="RALBase"
End
Begin OutputColumns
    Expression ="RALBase.clientID"
    Expression ="RALBase.SchID"
    Expression ="RALBase.RemitAddressID"
    Expression ="RALBase.ScheduleDate"
    Expression ="RALBase.Schedule"
    Expression ="RALBase.MLNo"
    Expression ="RALBase.MLDate"
    Expression ="RALBase.Lessee"
    Expression ="RALBase.LesseeAddress"
    Alias ="Paragraph1"
    Expression ="\"Reference is made to the Master Equipment Lease Agreement, dated \" & [mldate]"
        " & \", between \" & [Lessee] & \" as lessee (\" & Chr(34) & \"Lessee\" & Chr(34)"
        " & \") and Fleet Advantage, LLC as lessor (\" & Chr(34) & \"Lessor'') No. \" & ["
        "MLNo] & \" (\" & Chr(34) & \"Agreement of Lease\" & Chr(34) & \") and in particu"
        "lar to <b>Schedule No.\" & [Schedule] & \" (\" & [groupsIn] & \")</b> dated as o"
        "f \" & [ScheduleDate] & \", thereunder (together with the incorporated terms of "
        "the Agreement of Lease, \" & Chr(34) & \"Schedule\" & Chr(34) & \") true copies "
        "of which are attached hereto. The beneficial interest thereby created has been s"
        "old and assigned to<b> \" & [BankName] & \"</b> (\" & Chr(34) & \"Assignee\" & C"
        "hr(34) & \").  Accordingly, you are hereby notified to remit all sums due and pa"
        "yable under the Schedule (\" & Chr(34) & \"payments\" & Chr(34) & \" including w"
        "ithout limitation, amounts due under Sections 2, 4 and 5 of the Agreement of Lea"
        "se) as follows:<br />\""
    Expression ="RALBase.BankRemitAddress"
    Alias ="Paragraph2a"
    Expression ="\"<br />Basic Rent and all other payments should be made directly to Assignee or"
        " to such other location as Assignee may designate until you are advised otherwis"
        "e by Assignee.  Please cause Assignee to be listed as additional insured, in acc"
        "ordance with the Agreement of Lease, on your liability insurance policy and send"
        " a certificate to them at Assignee, c/o \""
    Alias ="Paragraph2b"
    Expression ="[BankAddressInline] & \". Please execute this letter in the space provided below"
        " and return it directly to Assignee, thereby acknowledging that  Lessee will mak"
        "e the payments as set forth above and in accordance with the Agreement of Lease,"
        " without claim, demand, offset, or defense of any kind whatsoever, that the Sche"
        "dule is in full force and effect and has not been amended, modified or altered a"
        "nd no default has occurred and is continuing thereunder, that Lessee has no clai"
        "ms, offsets, demands or defenses of any kind, nature or description with referen"
        "ce to any of the Lessor's obligations under the Schedule as of the date hereof a"
        "nd that Lessee will send notices thereunder to, and obtain any relevant consents"
        " from Assignee with a copy to Lessor.\""
    Alias ="Paragraph3a"
    Expression ="\"<br />Your execution of this letter also confirms the following to Assignee:  "
        "(i) the Schedule has been duly authorized, executed and delivered by Lessee; (ii"
        ") the Schedule is in full force and effect and is legal, valid, binding and enfo"
        "rceable against Lessee; (iii) the Lessee is not in default under the Schedule as"
        " the date hereof; \""
    Alias ="Paragraph3b"
    Expression ="\"(iv) that only the copy of the Schedule marked ''original'' is chattel paper f"
        "or purposes of transferring an interest in the Schedule and that any other copy "
        "of the Schedule is not the \" & Chr(34) & \"original\" & Chr(34) & \" for such p"
        "urposes; (v) Assignee is not responsible for the performance of any Lessor coven"
        "ant or agreement under the Schedule (except that Assignee agrees it will not int"
        "erfere with Lessee's quiet enjoyment of the Equipment so long as Lessee is not i"
        "n default); \""
    Alias ="Paragraph3c"
    Expression ="\"(vi) upon an Event of Default under the Schedule including but not limited to "
        "the failure of Lessee to make any and all payments required thereunder, Assignee"
        " may avail itself of any or all of the remedies provided for in the Schedule; an"
        "d (vii) there are no other agreements, amendments or riders which would modify a"
        "ny of the terms of the attached Agreement of Lease and Schedule.\""
    Alias ="Paragraph4"
    Expression ="\"<br />Upon and as of the date which the above-referenced assignment and sale t"
        "akes place, and as an inducement to Assignee to enter into same, Lessee agrees t"
        "hat the terms of the Schedule are as follows:  This letter also confirms that th"
        "e amounts and due dates of the Basic Rent payments under said Schedule are as fo"
        "llows:\""
    Alias ="Paragraph5"
    Expression ="\"It being understood, the Schedule shall consist of \" & [ConsistOf] & \" shall"
        " constitute a final independent and completed Lease.\015\012\015\012This letter "
        "may be executed in counterparts, which together shall constitute one and the sam"
        "e instrument.\""
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
        dbText "Name" ="Paragraph3b"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph1"
        dbInteger "ColumnWidth" ="5790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph3a"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.[BankRemitAddress]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph2a"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph2b"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph3c"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.LesseeAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.RemitAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.MLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.ScheduleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALBase.BankRemitAddress"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-487
    Top =102
    Right =806
    Bottom =904
    Left =-1
    Top =-1
    Right =1275
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =41
        Top =15
        Right =185
        Bottom =367
        Top =0
        Name ="RALBase"
        Name =""
    End
End

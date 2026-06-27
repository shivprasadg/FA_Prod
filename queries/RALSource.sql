SELECT
  RALBase.clientID,
  RALBase.SchID,
  RALBase.RemitAddressID,
  RALBase.ScheduleDate,
  RALBase.Schedule,
  RALBase.MLNo,
  RALBase.MLDate,
  RALBase.Lessee,
  RALBase.LesseeAddress,
  "Reference is made to the Master Equipment Lease Agreement, dated " & [mldate] & ", between " & [Lessee] & " as lessee (" & Chr(34)& "Lessee" & Chr(34)& ") and Fleet Advantage, LLC as lessor (" & Chr(34)& "Lessor'') No. " & [MLNo] & " (" & Chr(34)& "Agreement of Lease" & Chr(34)& ") and in particular to <b>Schedule No." & [Schedule] & " (" & [groupsIn] & ")</b> dated as of " & [ScheduleDate] & ", thereunder (together with the incorporated terms of the Agreement of Lease, " & Chr(34)& "Schedule" & Chr(34)& ") true copies of which are attached hereto. The beneficial interest thereby created has been sold and assigned to<b> " & [BankName] & "</b> (" & Chr(34)& "Assignee" & Chr(34)& ").  Accordingly, you are hereby notified to remit all sums due and payable under the Schedule (" & Chr(34)& "payments" & Chr(34)& " including without limitation, amounts due under Sections 2, 4 and 5 of the Agreement of Lease) as follows:<br />" AS Paragraph1,
  RALBase.BankRemitAddress,
  "<br />Basic Rent and all other payments should be made directly to Assignee or to such other location as Assignee may designate until you are advised otherwise by Assignee.  Please cause Assignee to be listed as additional insured, in accordance with the Agreement of Lease, on your liability insurance policy and send a certificate to them at Assignee, c/o " AS Paragraph2a,
  [BankAddressInline] & ". Please execute this letter in the space provided below and return it directly to Assignee, thereby acknowledging that  Lessee will make the payments as set forth above and in accordance with the Agreement of Lease, without claim, demand, offset, or defense of any kind whatsoever, that the Schedule is in full force and effect and has not been amended, modified or altered and no default has occurred and is continuing thereunder, that Lessee has no claims, offsets, demands or defenses of any kind, nature or description with reference to any of the Lessor's obligations under the Schedule as of the date hereof and that Lessee will send notices thereunder to, and obtain any relevant consents from Assignee with a copy to Lessor." AS Paragraph2b,
  "<br />Your execution of this letter also confirms the following to Assignee:  (i) the Schedule has been duly authorized, executed and delivered by Lessee; (ii) the Schedule is in full force and effect and is legal, valid, binding and enforceable against Lessee; (iii) the Lessee is not in default under the Schedule as the date hereof; " AS Paragraph3a,
  "(iv) that only the copy of the Schedule marked ''original'' is chattel paper for purposes of transferring an interest in the Schedule and that any other copy of the Schedule is not the " & Chr(34)& "original" & Chr(34)& " for such purposes; (v) Assignee is not responsible for the performance of any Lessor covenant or agreement under the Schedule (except that Assignee agrees it will not interfere with Lessee's quiet enjoyment of the Equipment so long as Lessee is not in default); " AS Paragraph3b,
  "(vi) upon an Event of Default under the Schedule including but not limited to the failure of Lessee to make any and all payments required thereunder, Assignee may avail itself of any or all of the remedies provided for in the Schedule; and (vii) there are no other agreements, amendments or riders which would modify any of the terms of the attached Agreement of Lease and Schedule." AS Paragraph3c,
  "<br />Upon and as of the date which the above-referenced assignment and sale takes place, and as an inducement to Assignee to enter into same, Lessee agrees that the terms of the Schedule are as follows:  This letter also confirms that the amounts and due dates of the Basic Rent payments under said Schedule are as follows:" AS Paragraph4,
  "It being understood, the Schedule shall consist of " & [ConsistOf] & " shall constitute a final independent and completed Lease.

This letter may be executed in counterparts, which together shall constitute one and the same instrument." AS Paragraph5
FROM
  RALBase;

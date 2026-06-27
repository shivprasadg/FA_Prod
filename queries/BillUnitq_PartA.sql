SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpID,
  Units.UnitID,
  IIf(
    IsNull([unitpdstartdate]),
    IIf(
      IsNull([unitinserv]),
      [unitacceptdate],
      [unitinserv]
    ),
    [unitpdstartdate]
  ) AS UBStart,
  SchGrp.BLCD AS unitmax,
  Units.UnitInServInd,
  Units.UnitStatus,
  Units.UnitDocAccpt,
  IIf(
    [sgrpnoslv] = No,
    IIf(
      [UBSTART] > [FORMS]![BillingSubForm]![BILLTO],
      [FORMS]![BillingSubForm]![BILLTO] + 1,
      IIf(
        [UBSTART] > [FORMS]![BillingSubForm]![BILLFROM],
        [UBSTART], [FORMS]![BillingSubForm]![BILLFROM]
      )
    ),
    IIf(
      [ubstart] >= [schcrexp] + 1
      And [ubstart] >= [forms]![BillingSubForm]![billfrom],
      [ubstart],
      IIf(
        [schcrexp] + 1 >= [forms]![BillingSubForm]![billfrom],
        [schcrexp] + 1, [forms]![BillingSubForm]![billfrom]
      )
    )
  ) AS UnitPDST,
  Units.UnitInServ
FROM
  (
    MstrLease
    INNER JOIN (
      (
        Schedule
        INNER JOIN Bills ON Schedule.SchID = Bills.BillSchRef
      )
      INNER JOIN SchGrp ON Bills.BillSchRef = SchGrp.SGrpSchID
    ) ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.UnitStatus)<> "b"
    )
    AND (
      (Units.UnitDocAccpt)= False
    )
  )
  OR (
    (
      (Units.UnitInServInd)= "a"
    )
    AND (
      (Units.UnitStatus)<> "b"
    )
    AND (
      (Units.UnitDocAccpt)= True
    )
    AND (
      (Units.UnitInServ) Is Not Null
    )
  )
  OR (
    (
      (Units.UnitInServInd)= "a"
    )
    AND (
      (Units.UnitStatus)<> "b"
    )
    AND (
      (Units.UnitDocAccpt)= True
    )
    AND (
      (Units.UnitInServ) Is Not Null
    )
    AND (
      (Units.UnitPDStartDate) Is Not Null
    )
  )
  OR (
    (
      (Units.UnitInServInd)= "a"
      And (Units.UnitInServInd) Is Not Null
    )
    AND (
      (Units.UnitStatus)= "b"
    )
  )
  OR (
    (
      (Units.UnitInServInd)= "p"
    )
    AND (
      (Units.UnitStatus)= "A"
    )
    AND (
      (Units.UnitPDStartDate) Is Not Null
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpID,
  Units.UnitID;

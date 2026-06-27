SELECT
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  Max(
    DateAdd(
      "m",
      1,
      IIf(
        IsNull([unitinserv]),
        IIf(
          IsNull([unitacceptdate]),
          IIf(
            IsNull([unitestdeldate]),
            [UNITESTDELDATE],
            [unitestdeldate]
          ),
          [unitacceptdate]
        ),
        [unitinserv]
      )
    )-1
  ) AS ProjSyn,
  MstrLease.MLOrig,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.LoanType,
  SchGrp.Assignee_Bank,
  SchGrp.FundDate,
  SchGrp.LoanBank,
  Sum(
    IIf(
      (
        [unitstatus] = "a"
        And [unitdocaccpt] = No
      )
      Or [unitinservind] = "a",
      1,
      0
    )
  ) AS AccptUnit,
  Sum(
    IIf(
      IsNull([UnitAcceptDate]),
      0,
      IIf(
        [unitdocaccpt] = Yes
        And [unitinservind] = "p",
        1, 0
      )
    )
  ) AS DocsAccpt,
  Max(Units.UnitBuildDate) AS MaxOfUnitBuildDate
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  SchGrp.SGrpSchID,
  SchGrp.SGrpID,
  MstrLease.MLOrig,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.LoanType,
  SchGrp.Assignee_Bank,
  SchGrp.FundDate,
  SchGrp.LoanBank
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (MstrLease.MLNo) Not Like "demo*"
    )
  )
  OR (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (MstrLease.MLNo) Not Like "demo*"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO;

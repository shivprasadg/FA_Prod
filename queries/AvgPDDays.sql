PARAMETERS [enter starting dt] DateTime,
[enter ending date] DateTime;
SELECT
  MstrLease.MLOrig,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.SGRPNoPD,
  Schedule.SchFSL,
  Schedule.SchSLVRec,
  SchGrp.SGRPINTPD,
  SchGrp.SGRPINTRT,
  SchGrp.SGrpNoSlv,
  Units.UnitPDStartDate,
  Units.UnitAcceptDate,
  Units.UnitInServ,
  IIf(
    IsNull([unitpdstartdate]),
    (
      IIf(
        IsNull([unitinserv]),
        [unitacceptdate],
        [unitinserv]
      )
    ),
    [unitpdstartdate]
  ) AS PDstart,
  IIf(
    (
      (
        DateDiff(
          "d",
          [BLCD],
          (
            IIf(
              IsNull([unitpdstartdate]),
              (
                IIf(
                  IsNull([unitinserv]),
                  [unitacceptdate],
                  [unitinserv]
                )
              ),
              [unitpdstartdate]
            )
          )
        )
      )*-1
    )< 0
    Or [sgrpnopd] =-1,
    0,
    (
      DateDiff(
        "d",
        [BLCD],
        (
          IIf(
            IsNull([unitpdstartdate]),
            (
              IIf(
                IsNull([unitinserv]),
                [unitacceptdate],
                [unitinserv]
              )
            ),
            [unitpdstartdate]
          )
        )
      )
    )*-1
  ) AS PDdays,
  (
    DateDiff(
      "d",
      [BLCD],
      (
        IIf(
          IsNull([unitpdstartdate]),
          (
            IIf(
              IsNull([unitinserv]),
              [unitacceptdate],
              [unitinserv]
            )
          ),
          [unitpdstartdate]
        )
      )
    )
  )*-1 AS nock
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (MstrLease.MLNo) Not Like "*demo*"
    )
    AND (
      (SchGrp.BLCD) Is Not Null
      And (SchGrp.BLCD)>= [enter starting dt]
      And (SchGrp.BLCD)<= [enter ending date]
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID;

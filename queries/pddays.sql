SELECT
  Clients.ClientShNm,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  Units.unitinserv,
  Units.unitacceptdate,
  Units.UnitPDStartDate,
  SchGrp.sgrpnopd,
  IIf(
    IsNull([unitpdstartdate]),
    IIf(
      IsNull([unitinserv]),
      [unitacceptdate],
      [unitinserv]
    ),
    [unitpdstartdate]
  ) AS PDStart,
  DateDiff(
    "d",
    IIf(
      IsNull([unitpdstartdate]),
      IIf(
        IsNull([unitinserv]),
        [unitacceptdate],
        [unitinserv]
      ),
      [unitpdstartdate]
    ),
    [unitacceptdate]
  ) AS DaysafteAccpt,
  IIf(
    DateDiff(
      "d",
      IIf(
        IsNull([unitpdstartdate]),
        IIf(
          IsNull([unitinserv]),
          [unitacceptdate],
          [unitinserv]
        ),
        [unitpdstartdate]
      ),
      [BLCD]
    )-1 < 0,
    0,
    DateDiff(
      "d",
      IIf(
        IsNull([unitpdstartdate]),
        IIf(
          IsNull([unitinserv]),
          [unitacceptdate],
          [unitinserv]
        ),
        [unitpdstartdate]
      ),
      [BLCD]
    )-1
  ) AS PDdays
FROM
  Clients
  INNER JOIN (
    (
      (
        MstrLease
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  ) ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (Clients.ClientShNm) Not Like "Demo*"
    )
    AND (
      (SchGrp.BLCD) Is Not Null
      And (SchGrp.BLCD)> #12/31/2010#
      And (SchGrp.BLCD)< #1/1/2012#
    )
    AND (
      (Schedule.SchNo)<> "1-ext"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO;
